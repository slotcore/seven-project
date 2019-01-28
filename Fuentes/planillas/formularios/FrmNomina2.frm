VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmNomina2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Nomina del Personal"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11760
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNomina2.frx":23EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
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
            Object.ToolTipText     =   "Exportar MsExcel"
            ImageIndex      =   12
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
      TabIndex        =   24
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
         TabIndex        =   27
         Top             =   375
         Width           =   11745
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   60
            TabIndex        =   28
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
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Apellidos y Nombres"
            Columns(2).DataField=   "nombres"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T.D."
            Columns(3).DataField=   "docabrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "N°.Doc."
            Columns(4).DataField=   "numdoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Categ."
            Columns(5).DataField=   "catnomcorto"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fch.Ingreso"
            Columns(6).DataField=   "fching1"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Cese"
            Columns(7).DataField=   "fchcese1"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Area"
            Columns(8).DataField=   "area"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Cargo"
            Columns(9).DataField=   "cargo"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Estado"
            Columns(10).DataField=   "estado"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Tipo Planilla"
            Columns(11).DataField=   "destippla"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1191"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1111"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4789"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4710"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1005"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=926"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=1826"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1746"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=1164"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1085"
            Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=1931"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1852"
            Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(39)=   "Column(7).Width=1773"
            Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=1693"
            Splits(0)._ColumnProps(42)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(44)=   "Column(8).Width=2090"
            Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2011"
            Splits(0)._ColumnProps(47)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(49)=   "Column(9).Width=2090"
            Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2011"
            Splits(0)._ColumnProps(52)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(54)=   "Column(10).Width=1482"
            Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=1402"
            Splits(0)._ColumnProps(57)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(59)=   "Column(11).Width=2858"
            Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=2778"
            Splits(0)._ColumnProps(62)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
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
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0,.bold=0,.fontsize=825"
            _StyleDefs(49)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(50)  =   ":id=28,.fontname=MS Sans Serif"
            _StyleDefs(51)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(52)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(55)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(56)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=78,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=46,.parent=13"
            _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
            _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
            _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(94)  =   "Named:id=33:Normal"
            _StyleDefs(95)  =   ":id=33,.parent=0"
            _StyleDefs(96)  =   "Named:id=34:Heading"
            _StyleDefs(97)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(98)  =   ":id=34,.wraptext=-1"
            _StyleDefs(99)  =   "Named:id=35:Footing"
            _StyleDefs(100) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(101) =   "Named:id=36:Selected"
            _StyleDefs(102) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(103) =   "Named:id=37:Caption"
            _StyleDefs(104) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(105) =   "Named:id=38:HighlightRow"
            _StyleDefs(106) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(107) =   "Named:id=39:EvenRow"
            _StyleDefs(108) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(109) =   "Named:id=40:OddRow"
            _StyleDefs(110) =   ":id=40,.parent=33"
            _StyleDefs(111) =   "Named:id=41:RecordSelector"
            _StyleDefs(112) =   ":id=41,.parent=34"
            _StyleDefs(113) =   "Named:id=42:FilterBar"
            _StyleDefs(114) =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Nómina del Personal"
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
            TabIndex        =   29
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   45
         TabIndex        =   25
         Top             =   375
         Width           =   11745
         Begin MSComDlg.CommonDialog cmm 
            Left            =   8610
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2910
            TabIndex        =   36
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   1365
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6165
            Left            =   -75
            TabIndex        =   30
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
            Caption         =   "   Datos Personales   |    Periodo Laboral    |    Derechohabiente   |   Datos de Trabajo   "
            Align           =   0
            CurrTab         =   1
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
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Height          =   5745
               Left            =   12645
               TabIndex        =   99
               Top             =   45
               Width           =   11610
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   14
                  Left            =   1650
                  Picture         =   "FrmNomina2.frx":277E
                  Style           =   1  'Graphical
                  TabIndex        =   160
                  Top             =   1320
                  Width           =   210
               End
               Begin VB.Frame FraHoras 
                  Caption         =   "Pago de Horas "
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
                  Height          =   630
                  Left            =   5700
                  TabIndex        =   156
                  Top             =   1630
                  Width           =   5325
                  Begin VB.TextBox txt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Height          =   285
                     Index           =   16
                     Left            =   3990
                     Locked          =   -1  'True
                     MaxLength       =   20
                     TabIndex        =   107
                     Tag             =   "null"
                     Text            =   "txt(16)"
                     Top             =   195
                     Width           =   1155
                  End
                  Begin VB.TextBox txt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Height          =   285
                     Index           =   15
                     Left            =   1500
                     Locked          =   -1  'True
                     MaxLength       =   20
                     TabIndex        =   106
                     Tag             =   "null"
                     Text            =   "txt(15)"
                     Top             =   195
                     Width           =   1155
                  End
                  Begin VB.Label lbl 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Horas Extras"
                     Height          =   195
                     Index           =   16
                     Left            =   2910
                     TabIndex        =   158
                     Top             =   270
                     Width           =   900
                  End
                  Begin VB.Label lbl 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Horas Normales"
                     Height          =   195
                     Index           =   15
                     Left            =   180
                     TabIndex        =   157
                     Top             =   300
                     Width           =   1125
                  End
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   13
                  Left            =   7485
                  Picture         =   "FrmNomina2.frx":28B0
                  Style           =   1  'Graphical
                  TabIndex        =   149
                  Top             =   1020
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   12
                  Left            =   7485
                  Picture         =   "FrmNomina2.frx":29E2
                  Style           =   1  'Graphical
                  TabIndex        =   148
                  Top             =   705
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   11
                  Left            =   1650
                  Picture         =   "FrmNomina2.frx":2B14
                  Style           =   1  'Graphical
                  TabIndex        =   143
                  Top             =   690
                  Width           =   210
               End
               Begin VB.Frame Frame10 
                  Caption         =   "[ Centro de Costos ]"
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
                  Height          =   2220
                  Left            =   165
                  TabIndex        =   114
                  Top             =   3135
                  Width           =   11160
                  Begin VB.TextBox txt_CenCos 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   7350
                     Locked          =   -1  'True
                     TabIndex        =   117
                     TabStop         =   0   'False
                     Text            =   "txt_CenCos"
                     Top             =   1770
                     Width           =   1050
                  End
                  Begin VB.Frame Frame9 
                     Height          =   1560
                     Left            =   9240
                     TabIndex        =   115
                     Top             =   195
                     Width           =   1680
                     Begin VB.CommandButton CmdCenCos 
                        Caption         =   "&Agregar Cento Costo"
                        Enabled         =   0   'False
                        Height          =   495
                        Index           =   0
                        Left            =   90
                        TabIndex        =   57
                        Top             =   330
                        Width           =   1470
                     End
                     Begin VB.CommandButton CmdCenCos 
                        Caption         =   "&Eliminar Cento Costo"
                        Enabled         =   0   'False
                        Height          =   495
                        Index           =   1
                        Left            =   90
                        TabIndex        =   116
                        Top             =   825
                        Width           =   1470
                     End
                  End
                  Begin VSFlex7Ctl.VSFlexGrid Fg3 
                     Height          =   1470
                     Left            =   150
                     TabIndex        =   58
                     Top             =   285
                     Width           =   8910
                     _cx             =   15716
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
                     FormatString    =   $"FrmNomina2.frx":2C46
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
                  Begin VB.Label Label2 
                     Caption         =   "Total ==>"
                     Height          =   195
                     Left            =   6435
                     TabIndex        =   118
                     Top             =   1815
                     Width           =   870
                  End
               End
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   3
                  Left            =   75
                  TabIndex        =   112
                  Top             =   165
                  Width           =   11460
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   9
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   8
                     X1              =   -30
                     X2              =   12000
                     Y1              =   390
                     Y2              =   390
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   6
                     X1              =   11445
                     X2              =   11445
                     Y1              =   -15
                     Y2              =   365
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Index           =   6
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   380
                  End
                  Begin VB.Label lbl_persona 
                     AutoSize        =   -1  'True
                     Caption         =   "lbl_persona(2)"
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
                     Index           =   2
                     Left            =   75
                     TabIndex        =   113
                     Top             =   75
                     Width           =   1215
                  End
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   10
                  Left            =   1650
                  Picture         =   "FrmNomina2.frx":2CE1
                  Style           =   1  'Graphical
                  TabIndex        =   145
                  Top             =   1005
                  Width           =   210
               End
               Begin VB.Frame FraEsPlanilla 
                  Caption         =   "¿Está en Planilla?"
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
                  Height          =   600
                  Left            =   195
                  TabIndex        =   103
                  Top             =   1630
                  Width           =   4920
                  Begin VB.TextBox txt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   12
                     Left            =   3495
                     Locked          =   -1  'True
                     MaxLength       =   20
                     TabIndex        =   105
                     Tag             =   "null"
                     Text            =   "txt(12)"
                     Top             =   195
                     Width           =   1155
                  End
                  Begin VB.OptionButton opt_esplanilla 
                     Caption         =   "No"
                     Height          =   285
                     Index           =   0
                     Left            =   555
                     TabIndex        =   54
                     Top             =   270
                     Value           =   -1  'True
                     Width           =   585
                  End
                  Begin VB.OptionButton opt_esplanilla 
                     Caption         =   "Si"
                     Height          =   285
                     Index           =   1
                     Left            =   1455
                     TabIndex        =   104
                     Top             =   255
                     Width           =   555
                  End
                  Begin VB.Label lbl 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sueldo Básico"
                     Height          =   195
                     Index           =   12
                     Left            =   2415
                     TabIndex        =   108
                     Top             =   285
                     Width           =   1020
                  End
               End
               Begin VB.Frame FraAsigFamiliar 
                  Caption         =   "¿Tiene Asignación Familiar?"
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
                  Height          =   600
                  Left            =   5700
                  TabIndex        =   101
                  Top             =   2430
                  Width           =   2745
                  Begin VB.OptionButton opt_asigfamiliar 
                     Caption         =   "Si"
                     Height          =   285
                     Index           =   1
                     Left            =   1455
                     TabIndex        =   102
                     Top             =   270
                     Width           =   555
                  End
                  Begin VB.OptionButton opt_asigfamiliar 
                     Caption         =   "No"
                     Height          =   285
                     Index           =   0
                     Left            =   555
                     TabIndex        =   55
                     Top             =   270
                     Value           =   -1  'True
                     Width           =   585
                  End
               End
               Begin VB.TextBox txt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   13
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   56
                  Tag             =   "null"
                  Text            =   "txt(13)"
                  Top             =   2700
                  Width           =   1155
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   10
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   51
                  Tag             =   "null"
                  Text            =   "txt_cb(10)"
                  Top             =   975
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   11
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   50
                  Tag             =   "null"
                  Text            =   "txt_cb(11)"
                  Top             =   660
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   13
                  Left            =   6975
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   53
                  Tag             =   "null"
                  Text            =   "txt_cb(13)"
                  Top             =   990
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   12
                  Left            =   6975
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   52
                  Tag             =   "null"
                  Text            =   "txt_cb(12)"
                  Top             =   675
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   14
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   161
                  Tag             =   "null"
                  Text            =   "txt_cb(11)"
                  Top             =   1290
                  Width           =   765
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(14)"
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
                  Index           =   14
                  Left            =   3300
                  TabIndex        =   164
                  Top             =   1260
                  Visible         =   0   'False
                  Width           =   1005
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo Planilla"
                  Height          =   255
                  Index           =   14
                  Left            =   195
                  TabIndex        =   163
                  Top             =   1350
                  Width           =   855
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(14)"
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
                  Index           =   14
                  Left            =   1890
                  TabIndex        =   162
                  Top             =   1290
                  Width           =   3660
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(13)"
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
                  Index           =   13
                  Left            =   9165
                  TabIndex        =   153
                  Top             =   990
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Boleta de Pago"
                  Height          =   195
                  Index           =   13
                  Left            =   5715
                  TabIndex        =   152
                  Top             =   1065
                  Width           =   1095
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lugar de Trabajo"
                  Height          =   195
                  Index           =   12
                  Left            =   5715
                  TabIndex        =   151
                  Top             =   750
                  Width           =   1215
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(12)"
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
                  Index           =   12
                  Left            =   9165
                  TabIndex        =   150
                  Top             =   675
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(11)"
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
                  Index           =   11
                  Left            =   3330
                  TabIndex        =   146
                  Top             =   660
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Area"
                  Height          =   195
                  Index           =   11
                  Left            =   195
                  TabIndex        =   144
                  Top             =   735
                  Width           =   330
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cargo"
                  Height          =   195
                  Index           =   10
                  Left            =   195
                  TabIndex        =   111
                  Top             =   1050
                  Width           =   420
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(10)"
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
                  Index           =   10
                  Left            =   3330
                  TabIndex        =   109
                  Top             =   975
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bonificación"
                  Height          =   195
                  Index           =   13
                  Left            =   195
                  TabIndex        =   100
                  Top             =   2730
                  Width           =   870
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(10)"
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
                  Index           =   10
                  Left            =   1890
                  TabIndex        =   110
                  Top             =   975
                  Width           =   3660
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(11)"
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
                  Index           =   11
                  Left            =   1890
                  TabIndex        =   147
                  Top             =   660
                  Width           =   3660
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(12)"
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
                  Index           =   12
                  Left            =   7725
                  TabIndex        =   155
                  Top             =   675
                  Width           =   3780
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(13)"
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
                  Index           =   13
                  Left            =   7725
                  TabIndex        =   154
                  Top             =   990
                  Width           =   3780
               End
            End
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   5745
               Left            =   12345
               TabIndex        =   91
               Top             =   45
               Width           =   11610
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   12
                  Left            =   75
                  TabIndex        =   96
                  Top             =   165
                  Width           =   11460
                  Begin VB.Label lbl_persona 
                     AutoSize        =   -1  'True
                     Caption         =   "lbl_persona(0)"
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
                     Index           =   0
                     Left            =   75
                     TabIndex        =   97
                     Top             =   75
                     Width           =   1215
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
                     X1              =   11445
                     X2              =   11445
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
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   585
                  Index           =   1
                  Left            =   75
                  TabIndex        =   92
                  Top             =   4815
                  Width           =   11460
                  Begin VB.CommandButton CmdDeHab 
                     Caption         =   "Eliminar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   2
                     Left            =   3525
                     TabIndex        =   95
                     Top             =   75
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdDeHab 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   0
                     Left            =   120
                     TabIndex        =   94
                     Top             =   75
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdDeHab 
                     Caption         =   "Modificar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   1
                     Left            =   1560
                     TabIndex        =   93
                     Top             =   75
                     Width           =   1395
                  End
                  Begin VB.Line Line4 
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
                     Index           =   4
                     X1              =   11445
                     X2              =   11445
                     Y1              =   0
                     Y2              =   985
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   4
                     X1              =   -30
                     X2              =   12000
                     Y1              =   570
                     Y2              =   570
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   5
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   4035
                  Left            =   75
                  TabIndex        =   98
                  Top             =   675
                  Width           =   11460
                  _cx             =   20214
                  _cy             =   7117
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
                  FormatString    =   $"FrmNomina2.frx":2E13
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
            Begin VB.Frame Frame7 
               BorderStyle     =   0  'None
               Height          =   5745
               Left            =   45
               TabIndex        =   82
               Top             =   45
               Width           =   11610
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   2
                  Left            =   75
                  TabIndex        =   88
                  Top             =   165
                  Width           =   11460
                  Begin VB.Label lbl_persona 
                     AutoSize        =   -1  'True
                     Caption         =   "lbl_persona(1)"
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
                     Index           =   1
                     Left            =   75
                     TabIndex        =   89
                     Top             =   75
                     Width           =   1215
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Index           =   5
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   380
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   5
                     X1              =   11445
                     X2              =   11445
                     Y1              =   -15
                     Y2              =   365
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   7
                     X1              =   -30
                     X2              =   12000
                     Y1              =   390
                     Y2              =   390
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   6
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
               End
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   585
                  Index           =   0
                  Left            =   75
                  TabIndex        =   83
                  Top             =   4815
                  Width           =   11460
                  Begin VB.CommandButton CmdInfCat 
                     Caption         =   "Ver Información Detallada"
                     Enabled         =   0   'False
                     Height          =   435
                     Left            =   8970
                     TabIndex        =   86
                     Top             =   75
                     Width           =   2235
                  End
                  Begin VB.CommandButton CmdPerLab 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   1
                     Left            =   1560
                     TabIndex        =   85
                     Top             =   75
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdPerLab 
                     Caption         =   "&Agregar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   0
                     Left            =   120
                     TabIndex        =   84
                     Top             =   75
                     Width           =   1395
                  End
                  Begin VB.Label lbl_categoria 
                     Alignment       =   1  'Right Justify
                     Caption         =   "lbl_categoria"
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
                     Left            =   4620
                     TabIndex        =   87
                     Top             =   270
                     Width           =   4245
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
                     X1              =   11445
                     X2              =   11445
                     Y1              =   0
                     Y2              =   985
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
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   4035
                  Left            =   75
                  TabIndex        =   90
                  Top             =   675
                  Width           =   11460
                  _cx             =   20214
                  _cy             =   7117
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
                  FormatString    =   $"FrmNomina2.frx":2F02
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5745
               Left            =   -12255
               TabIndex        =   31
               Top             =   45
               Width           =   11610
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   14
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   22
                  Tag             =   "null"
                  Text            =   "txt(14)"
                  Top             =   5370
                  Width           =   2280
               End
               Begin VB.Frame Frame5 
                  BorderStyle     =   0  'None
                  Caption         =   "2"
                  Height          =   3645
                  Left            =   8340
                  TabIndex        =   80
                  Top             =   105
                  Width           =   3165
                  Begin VB.CommandButton CmdFoto 
                     Caption         =   "Eliminar Imagen"
                     Enabled         =   0   'False
                     Height          =   375
                     Index           =   1
                     Left            =   1650
                     TabIndex        =   137
                     Top             =   3180
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdFoto 
                     Caption         =   "Agregar Imagen"
                     Enabled         =   0   'False
                     Height          =   375
                     Index           =   0
                     Left            =   150
                     TabIndex        =   136
                     Top             =   3180
                     Width           =   1395
                  End
                  Begin VB.PictureBox pic 
                     AutoSize        =   -1  'True
                     Height          =   2790
                     Index           =   0
                     Left            =   120
                     ScaleHeight     =   2730
                     ScaleWidth      =   2850
                     TabIndex        =   81
                     Top             =   315
                     Width           =   2910
                  End
                  Begin VB.Label lbl_pic 
                     Alignment       =   2  'Center
                     Caption         =   "Fotografía"
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
                     Index           =   0
                     Left            =   105
                     TabIndex        =   141
                     Top             =   75
                     Width           =   2940
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H80000003&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   3150
                     X2              =   3150
                     Y1              =   30
                     Y2              =   4260
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H80000005&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   15
                     X2              =   15
                     Y1              =   15
                     Y2              =   4245
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H80000003&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   45
                     X2              =   3675
                     Y1              =   3630
                     Y2              =   3630
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H80000005&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   0
                     X2              =   3630
                     Y1              =   15
                     Y2              =   15
                  End
               End
               Begin VB.Frame Frame6 
                  BorderStyle     =   0  'None
                  Caption         =   "&H80000005&"
                  Height          =   1800
                  Left            =   8340
                  TabIndex        =   79
                  Top             =   3795
                  Width           =   3165
                  Begin VB.PictureBox pic 
                     Height          =   990
                     Index           =   1
                     Left            =   105
                     ScaleHeight     =   930
                     ScaleWidth      =   2850
                     TabIndex        =   140
                     Top             =   315
                     Width           =   2910
                  End
                  Begin VB.CommandButton CmdFirma 
                     Caption         =   "Eliminar Firma"
                     Enabled         =   0   'False
                     Height          =   375
                     Index           =   1
                     Left            =   1620
                     TabIndex        =   139
                     Top             =   1365
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdFirma 
                     Caption         =   "Agregar Firma"
                     Enabled         =   0   'False
                     Height          =   375
                     Index           =   0
                     Left            =   120
                     TabIndex        =   138
                     Top             =   1365
                     Width           =   1395
                  End
                  Begin VB.Label lbl_pic 
                     Alignment       =   2  'Center
                     Caption         =   "Firma"
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
                     Index           =   1
                     Left            =   120
                     TabIndex        =   142
                     Top             =   75
                     Width           =   2940
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H80000005&
                     BorderWidth     =   2
                     Index           =   2
                     X1              =   0
                     X2              =   3630
                     Y1              =   15
                     Y2              =   15
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H80000003&
                     BorderWidth     =   2
                     Index           =   3
                     X1              =   -15
                     X2              =   3615
                     Y1              =   1785
                     Y2              =   1785
                  End
                  Begin VB.Line Line3 
                     BorderColor     =   &H80000003&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   3150
                     X2              =   3150
                     Y1              =   15
                     Y2              =   1815
                  End
                  Begin VB.Line Line3 
                     BorderColor     =   &H80000005&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   15
                     X2              =   15
                     Y1              =   30
                     Y2              =   1830
                  End
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   5
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":300A
                  Style           =   1  'Graphical
                  TabIndex        =   74
                  Top             =   2805
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   4
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":313C
                  Style           =   1  'Graphical
                  TabIndex        =   72
                  Top             =   2460
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   3
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":326E
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  Top             =   2130
                  Width           =   210
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   11
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   20
                  Tag             =   "null"
                  Text            =   "txt(11)"
                  Top             =   4725
                  Width           =   6675
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   10
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   19
                  Tag             =   "null"
                  Text            =   "txt(10)"
                  Top             =   4410
                  Width           =   3795
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   8
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":33A0
                  Style           =   1  'Graphical
                  TabIndex        =   63
                  Top             =   4110
                  Width           =   210
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   9
                  Left            =   6780
                  Locked          =   -1  'True
                  MaxLength       =   4
                  TabIndex        =   17
                  Tag             =   "null"
                  Text            =   "txt(9)"
                  Top             =   3765
                  Width           =   1035
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   8
                  Left            =   5055
                  Locked          =   -1  'True
                  MaxLength       =   4
                  TabIndex        =   16
                  Tag             =   "null"
                  Text            =   "txt(8)"
                  Top             =   3765
                  Width           =   1035
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   7
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   15
                  Tag             =   "null"
                  Text            =   "txt(7)"
                  Top             =   3765
                  Width           =   3180
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   7
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":34D2
                  Style           =   1  'Graphical
                  TabIndex        =   59
                  Top             =   3465
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   6
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":3604
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  Top             =   3135
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   9
                  Left            =   1995
                  Picture         =   "FrmNomina2.frx":3736
                  Style           =   1  'Graphical
                  TabIndex        =   46
                  Top             =   5070
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   2
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":3868
                  Style           =   1  'Graphical
                  TabIndex        =   44
                  Top             =   1485
                  Width           =   210
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   2
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   7
                  Text            =   "txt_cb(2)"
                  Top             =   1455
                  Width           =   765
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   6
                  Left            =   4110
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   9
                  Tag             =   "null"
                  Text            =   "txt(6)"
                  Top             =   1785
                  Width           =   3705
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   5
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   8
                  Tag             =   "null"
                  Text            =   "txt(5)"
                  Top             =   1785
                  Width           =   1680
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   1
                  Left            =   4740
                  Picture         =   "FrmNomina2.frx":399A
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  ToolTipText     =   "Seleccione el Sexo"
                  Top             =   1155
                  Width           =   210
               End
               Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
                  Height          =   300
                  Index           =   0
                  Left            =   1140
                  TabIndex        =   5
                  Top             =   1125
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
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   4
                  Left            =   5790
                  MaxLength       =   15
                  TabIndex        =   4
                  Tag             =   "null"
                  Text            =   "txt(4)"
                  Top             =   795
                  Width           =   1485
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   3
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   2
                  Text            =   "txt(3)"
                  Top             =   480
                  Width           =   6675
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   2
                  Left            =   5055
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   1
                  Tag             =   "null"
                  Text            =   "txt(2)"
                  Top             =   165
                  Width           =   2760
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   1
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   0
                  Text            =   "txt(1)"
                  Top             =   165
                  Width           =   2715
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   0
                  Left            =   1680
                  Picture         =   "FrmNomina2.frx":3ACC
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  ToolTipText     =   "Seleccione el Tipo de Documento"
                  Top             =   825
                  Width           =   210
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   0
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   3
                  Text            =   "txt_cb(0)"
                  Top             =   795
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   1
                  Left            =   4215
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   6
                  Text            =   "txt_cb(1)"
                  ToolTipText     =   "Ingrese el Sexo (1:Masculino, 2:Femenino)"
                  Top             =   1125
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   9
                  Left            =   1485
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   21
                  Tag             =   "null"
                  Text            =   "txt_cb(9)"
                  Top             =   5040
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   6
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   13
                  Tag             =   "null"
                  Text            =   "txt_cb(6)"
                  Top             =   3105
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   7
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   14
                  Tag             =   "null"
                  Text            =   "txt_cb(7)"
                  Top             =   3435
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   8
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   18
                  Tag             =   "null"
                  Text            =   "txt_cb(8)"
                  Top             =   4080
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   3
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   10
                  Tag             =   "null"
                  Text            =   "txt_cb(3)"
                  Top             =   2100
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   4
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   11
                  Tag             =   "null"
                  Text            =   "txt_cb(4)"
                  Top             =   2430
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   5
                  Left            =   1140
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   12
                  Tag             =   "null"
                  Text            =   "txt_cb(5)"
                  Top             =   2775
                  Width           =   765
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo Doc."
                  Height          =   195
                  Index           =   0
                  Left            =   105
                  TabIndex        =   135
                  Top             =   905
                  Width           =   705
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ap. Paterno"
                  Height          =   195
                  Index           =   1
                  Left            =   105
                  TabIndex        =   134
                  Top             =   255
                  Width           =   840
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombres"
                  Height          =   195
                  Index           =   3
                  Left            =   105
                  TabIndex        =   133
                  Top             =   580
                  Width           =   630
               End
               Begin VB.Label lbl_fecha 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nacimiento"
                  Height          =   195
                  Index           =   0
                  Left            =   105
                  TabIndex        =   132
                  Top             =   1230
                  Width           =   795
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Teléfono"
                  Height          =   195
                  Index           =   5
                  Left            =   105
                  TabIndex        =   131
                  Top             =   1880
                  Width           =   630
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nacionalidad"
                  Height          =   195
                  Index           =   2
                  Left            =   105
                  TabIndex        =   130
                  Top             =   1555
                  Width           =   930
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ESSALUD + Vida "
                  Height          =   195
                  Index           =   9
                  Left            =   105
                  TabIndex        =   129
                  Top             =   5130
                  Width           =   1290
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Domiciliado"
                  Height          =   195
                  Index           =   6
                  Left            =   105
                  TabIndex        =   128
                  Top             =   3180
                  Width           =   810
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo Vía"
                  Height          =   195
                  Index           =   7
                  Left            =   105
                  TabIndex        =   127
                  Top             =   3505
                  Width           =   615
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre Vía"
                  Height          =   195
                  Index           =   7
                  Left            =   105
                  TabIndex        =   126
                  Top             =   3830
                  Width           =   855
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo Zona"
                  Height          =   195
                  Index           =   8
                  Left            =   105
                  TabIndex        =   125
                  Top             =   4155
                  Width           =   735
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre Zona"
                  Height          =   195
                  Index           =   10
                  Left            =   105
                  TabIndex        =   124
                  Top             =   4480
                  Width           =   975
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Referencia"
                  Height          =   195
                  Index           =   11
                  Left            =   105
                  TabIndex        =   123
                  Top             =   4805
                  Width           =   780
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Departamento"
                  Height          =   195
                  Index           =   3
                  Left            =   105
                  TabIndex        =   122
                  Top             =   2205
                  Width           =   1005
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Provincia"
                  Height          =   195
                  Index           =   4
                  Left            =   105
                  TabIndex        =   121
                  Top             =   2530
                  Width           =   660
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Distrito"
                  Height          =   195
                  Index           =   5
                  Left            =   105
                  TabIndex        =   120
                  Top             =   2855
                  Width           =   480
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nº ESSALUD"
                  Height          =   195
                  Index           =   14
                  Left            =   105
                  TabIndex        =   119
                  Top             =   5460
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(5)"
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
                  Index           =   5
                  Left            =   3675
                  TabIndex        =   78
                  Top             =   2775
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(4)"
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
                  Index           =   4
                  Left            =   3675
                  TabIndex        =   77
                  Top             =   2430
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(3)"
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
                  Index           =   3
                  Left            =   3675
                  TabIndex        =   76
                  Top             =   2100
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(5)"
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
                  Index           =   5
                  Left            =   1905
                  TabIndex        =   75
                  Top             =   2775
                  Width           =   3045
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(4)"
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
                  Index           =   4
                  Left            =   1905
                  TabIndex        =   73
                  Top             =   2430
                  Width           =   3045
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(3)"
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
                  Index           =   3
                  Left            =   1905
                  TabIndex        =   71
                  Top             =   2100
                  Width           =   3045
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(8)"
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
                  Index           =   8
                  Left            =   3675
                  TabIndex        =   69
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(7)"
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
                  Index           =   7
                  Left            =   3675
                  TabIndex        =   68
                  Top             =   3435
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod6)"
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
                  Index           =   6
                  Left            =   3675
                  TabIndex        =   67
                  Top             =   3105
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(9)"
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
                  Index           =   9
                  Left            =   3675
                  TabIndex        =   66
                  Top             =   5040
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(2)"
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
                  Index           =   2
                  Left            =   3675
                  TabIndex        =   65
                  Top             =   1455
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(8)"
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
                  Index           =   8
                  Left            =   1905
                  TabIndex        =   64
                  Top             =   4080
                  Width           =   3045
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Interior"
                  Height          =   195
                  Index           =   9
                  Left            =   6210
                  TabIndex        =   62
                  Top             =   3830
                  Width           =   480
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Número"
                  Height          =   195
                  Index           =   8
                  Left            =   4410
                  TabIndex        =   61
                  Top             =   3830
                  Width           =   555
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(7)"
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
                  Index           =   7
                  Left            =   1905
                  TabIndex        =   60
                  Top             =   3435
                  Width           =   3045
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(6)"
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
                  Index           =   6
                  Left            =   1905
                  TabIndex        =   49
                  Top             =   3105
                  Width           =   3045
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(9)"
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
                  Index           =   9
                  Left            =   2235
                  TabIndex        =   47
                  Top             =   5040
                  Width           =   2715
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(2)"
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
                  Index           =   2
                  Left            =   1905
                  TabIndex        =   45
                  Top             =   1455
                  Width           =   3045
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-Mail"
                  Height          =   195
                  Index           =   6
                  Left            =   3540
                  TabIndex        =   43
                  Top             =   1880
                  Width           =   435
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
                  Left            =   6360
                  TabIndex        =   42
                  Top             =   1125
                  Visible         =   0   'False
                  Width           =   975
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
                  Left            =   4995
                  TabIndex        =   41
                  Top             =   1125
                  Width           =   2685
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sexo"
                  Height          =   195
                  Index           =   1
                  Left            =   3720
                  TabIndex        =   40
                  Top             =   1230
                  Width           =   360
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Número"
                  Height          =   195
                  Index           =   4
                  Left            =   5130
                  TabIndex        =   38
                  Top             =   905
                  Width           =   555
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ap. Materno"
                  Height          =   195
                  Index           =   2
                  Left            =   4020
                  TabIndex        =   35
                  Top             =   255
                  Width           =   870
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
                  Left            =   3675
                  TabIndex        =   33
                  Top             =   795
                  Visible         =   0   'False
                  Width           =   975
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
                  Left            =   1905
                  TabIndex        =   34
                  Top             =   795
                  Width           =   3045
               End
            End
         End
         Begin VB.Label LblCodigoEmp 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCodigoEmp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10080
            TabIndex        =   159
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   9570
            TabIndex        =   37
            Top             =   75
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Datos del Personal"
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
            TabIndex        =   26
            Top             =   60
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmNomina2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim Quehace As Integer
Dim Agregando As Boolean

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim nSQL As String
Dim SINCPLANILLA_ As Boolean


Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle del Personal"
    Quehace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    
    TabOne2.TabEnabled(2) = True
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    
    TabOne1.CurrTab = 0
End Sub


Private Sub CmdDeHab_Click(Index As Integer)
    On Error GoTo error
    Select Case Index
        Case 0 '--agregar
            FrmDerechohabiente.pRecibeLink 1
            FrmDerechohabiente.Show 1
        Case 1 '--modificar
            If Fg1.Row < 1 Then Exit Sub
            If Fg1.Rows = 1 Then
                MsgBox "No hay registro", vbExclamation, xTitulo
                Exit Sub
            End If
            FrmDerechohabiente.pRecibeLink 2, NulosN(Fg1.TextMatrix(Fg1.Row, 1))
            FrmDerechohabiente.Show 1
        Case 2 '--eliminar
            If Fg1.Row < 1 Then Exit Sub
            If Fg1.Rows = 1 Then
                MsgBox "No hay registro", vbExclamation, xTitulo
                Exit Sub
            End If
            If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            '--ELIMINADO REGISTRO
            xCon.Execute "DELETE FROM pla_derechohab WHERE idemp=" & RstFrm.Fields("id") & " and corr = " & Fg1.TextMatrix(Fg1.Row, 1) & " ;"
            pCargarDatosDerechoHabiente
            MsgBox "El registro se eliminó correctamente", vbInformation, xTitulo
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "CmdDeHab_Click (" & Index & ")"
End Sub



Private Sub CmdFoto_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pAddFoto 1, pic(0)
        Case 1 '--eliminar
            
    End Select
End Sub

Private Sub CmdFirma_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pAddFoto 2, pic(1)
        Case 1 '--eliminar
            
    End Select

End Sub

Private Sub CmdInfCat_Click()
    '--peridos laborables
    If Fg2.Row < 1 Then Exit Sub
    
    fg2_CellButtonClick Fg2.Row, 7
        
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

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 7 Then KeyAscii = 0
End Sub

Private Sub Fg2_RowColChange()
    If Fg2.Row < 1 Then Exit Sub
    If NulosN(Fg2.Cell(flexcpText, Fg2.Row, 2)) = 0 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    RST_Busq RstTmp, "SELECT mae_categoria.descripcion FROM mae_categoria WHERE (((mae_categoria.id)=" & Fg2.Cell(flexcpText, Fg2.Row, 2) & "))", xCon
    If RstTmp.RecordCount <> 0 Then
        lbl_categoria.Caption = RstTmp.Fields(0)
    End If
    Set RstTmp = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
   
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
            
        pConfigurarGrilla
        pCargarGrid
    End If
    
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Quehace = 3
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    Dim xRs As New ADODB.Recordset
    
    Dg1.Columns("fching1").NumberFormat = "DD/MM/YYYY"
    Dg1.Columns("fchcese1").NumberFormat = "DD/MM/YYYY"
    
    TabOne1.CurrTab = 0
    Fg1.ColWidth(1) = 0
    
    nSQL = "SELECT mae_empresa.sincpla " _
            + vbCr + "FROM mae_empresa;"
    
    RST_Busq xRs, nSQL, xCon
    
    If xRs.State = 0 Then SINCPLANILLA_ = False: Exit Sub
    If xRs.RecordCount = 0 Then SINCPLANILLA_ = False: Exit Sub
    
    SINCPLANILLA_ = xRs("sincpla")
End Sub


Sub Blanquea()
    lbl_categoria.Caption = ""
    LimpiaText txt
    LimpiaText txtfecha
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
    LimpiaText lbl_persona
    txt_CenCos.Text = ""
    '--limpiando las fotos
    pic(0).Picture = LoadPicture("")
    pic(1).Picture = LoadPicture("")
    Fg2.Rows = 1
    Fg1.Rows = 1
    Fg3.Rows = 1
    LblCodigoEmp.Caption = ""
    
End Sub

Sub Bloquea(band As Boolean)
    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    habilitar CmdFoto, band
    habilitar CmdFirma, band
    habilitar CmdDeHab, band
    habilitar CmdPerLab, band
    habilitar_Locked txtfecha, Not band
    CmdInfCat.Enabled = Not band
        
    habilitar CmdCenCos, band
    FraEsPlanilla.Enabled = band
    FraAsigFamiliar.Enabled = band
    FraHoras.Enabled = band
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Quehace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg1.DataSource = Nothing
    Set RstFrm = Nothing
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If Quehace <> 1 Then MuestraSegundoTab
    End If
End Sub

Function AbrirConecciones(Ruta As String) As ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCone As ADODB.Connection
    
    xFun.F_BASEDATOS = Ruta
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCone = xFun.AbrirConeccion
    Set xFun = Nothing
    Set AbrirConecciones = xCone
End Function

Private Sub procesarRestoEmpresas(TIPO_ As Double)
    '******************
    'TIPO: 1 GRABAR
    'TIPO: 2 MODIFICAR
    '******************
    
    Dim xConTemp As New ADODB.Connection
    Dim xConAux As New ADODB.Connection
    Dim rstTemp As New ADODB.Recordset
    Dim A As Integer
    Dim xIndex As Integer
    Dim RUTA_ As String
    
    Set xConTemp = AbrirConecciones(AP_RUTABD + "data.mdb")
    
    nSQL = "SELECT mae_empresa.* " _
        + vbCr + "FROM mae_empresa " _
        + vbCr + "WHERE (((mae_empresa.anotra)=" & AnoTra & ") AND ((mae_empresa.activo)=-1) " _
                                                    & "AND ((mae_empresa.numruc)<>'" & NumRuc & "'))"
    
    RST_Busq rstTemp, nSQL, xConTemp
    
    If rstTemp.RecordCount = 0 Then Exit Sub
    
    rstTemp.MoveFirst
    
    For A = 1 To rstTemp.RecordCount
        Set xConAux = Nothing
        
        RUTA_ = AP_RUTABD + Trim(rstTemp("ruta"))
        Set xConAux = AbrirConecciones(RUTA_)
        
        Select Case TIPO_
            Case 1
                If Not Grabar(xConAux, False) Then GoTo SALIR_
                
            Case 2
                Eliminar xConAux, False
        
        End Select
                
        rstTemp.MoveNext
        If rstTemp.EOF Then Exit For
    Next A
    
    Set rstTemp = Nothing
    Set xConAux = Nothing
    Set xConTemp = Nothing
SALIR_:

End Sub

Function Grabar(ByRef XCON_ As ADODB.Connection, Optional PRIMERO_ As Boolean = True) As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "No ha especificado el tipo de documento de identidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If txt(4).Text = "" Then
        MsgBox "No ha especificado el numero de documento de identidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txt(4).SetFocus
        Exit Function
    Else
        Dim Rst As New ADODB.Recordset
        
        RST_Busq Rst, "SELECT * FROM pla_empleados WHERE idtipdoc = " & NulosN(txt_cb(0).Text) & " AND numdoc = '" & txt(4).Text & "'", XCON_
        If Quehace = 1 Then
            If Rst.RecordCount <> 0 Then
                MsgBox "El tipo de documento " & lbl_cb(0).Caption & " con Nº " & txt(4).Text & " ya fue registrado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                txt(4).SetFocus
                Exit Function
            End If
        End If
        If Quehace = 2 Then
            If Rst.RecordCount > 1 Then
                MsgBox "El tipo de documento " & lbl_cb(0).Caption & " con Nº " & txt(4).Text & " ya fue registrado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                txt(4).SetFocus
                Exit Function
            End If
        End If
    End If
    '---------
    If Quehace = 1 Then
        DefinirCodigo
    Else
        '--verificar si el personal no tiene
        If NulosC(RstFrm("codigo")) = "" Then DefinirCodigo
    End If
    DoEvents
    '---------
    If PRIMERO_ Then
        If MsgBox("Seguro desea " + IIf(Quehace = 1, "Grabar", "Modificar") + " al Personal ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    End If

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstPerLab As New ADODB.Recordset
    Dim RstCenCos As New ADODB.Recordset '--relacionado con el centro de costo
    Dim RstBus As New ADODB.Recordset
    Dim nSQL As String
    Dim xId As Double
    Dim A&

On Error GoTo LaCague

    XCON_.BeginTrans

    If Quehace = 1 Then

        xId = HallaCodigoTabla("pla_empleados", XCON_, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_empleados", XCON_

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pla_empleados WHERE id = " & xId & "", XCON_
        '--eliminar el periodo laboral
        XCON_.Execute "Delete from pla_periodolaboral where idemp =  " & xId & " ;"
        '--eliminar los conceptos relacionado con el centro de costo
        XCON_.Execute "DELETE FROM pla_empleadoscos WHERE idemp = " & xId & ";"
        '--eliminar las imagenes relacionadas
        XCON_.Execute "DELETE FROM pla_empleadosimg WHERE idemp = " & xId & ";"
        
    End If
    
    txt(0).Text = xId
    
    '--almacenar el id temporalmente
    mIdRegistro = xId
    '---
    RST_Busq RstPerLab, "SELECT TOP 1 * FROM pla_periodolaboral ; ", XCON_
    RST_Busq RstCenCos, "SELECT TOP 1 * FROM pla_empleadoscos", XCON_
    
    RstCab("codigo") = LblCodigoEmp.Caption
    
    RstCab("idtipdoc") = NulosN(txt_cb(0).Text)
    RstCab("numdoc") = Trim(txt(4).Text)
    
    RstCab("apepat") = Trim(txt(1).Text)
    RstCab("apemat") = Trim(txt(2).Text)
    RstCab("nom") = Trim(txt(3).Text)
    
    If IsDate(txtfecha(0).Valor) = True Then RstCab("fchnac") = CDate(txtfecha(0).Valor)
    
    RstCab("idsex") = NulosN(lbl_cod(1).Caption)
    RstCab("idnac") = NulosN(lbl_cod(2).Caption)
    RstCab("numtel") = Trim(txt(5).Text)
    RstCab("email") = Trim(txt(6).Text)
    RstCab("indessalud") = NulosN(lbl_cod(9).Caption)
    RstCab("numessalud") = NulosC(txt(14).Text)
    RstCab("inddomi") = NulosN(lbl_cod(6).Caption)
    RstCab("idtipvia") = NulosN(lbl_cod(7).Caption)
    RstCab("nomvia") = Trim(txt(7).Text)
    RstCab("numvia") = Trim(txt(8).Text)

    RstCab("intvia") = Trim(txt(9).Text)
    RstCab("idtipzon") = NulosN(lbl_cod(8).Caption)

    RstCab("nomzon") = Trim(txt(10).Text)

    RstCab("refdom") = Trim(txt(11).Text)
    
    RstCab("iddep") = NulosN(lbl_cod(3).Caption)
    RstCab("idpro") = NulosN(lbl_cod(4).Caption)
    RstCab("iddis") = NulosN(lbl_cod(5).Caption)
    
    RstCab("idarea") = NulosN(lbl_cod(11).Caption)
    RstCab("idcargo") = NulosN(lbl_cod(10).Caption)
    
    RstCab("idlugtra") = NulosN(lbl_cod(12).Caption)
    RstCab("idbolpag") = NulosN(lbl_cod(13).Caption)
    
    ' *************************************************************
    RstCab("idtippla") = NulosN(lbl_cod(14).Caption)
    ' *************************************************************
    
    If opt_esplanilla(0).Value = True Then RstCab("aplanilla") = 0
    If opt_esplanilla(1).Value = True Then RstCab("aplanilla") = -1
    
    If opt_asigfamiliar(0).Value = True Then RstCab("asigfam") = 0
    If opt_asigfamiliar(1).Value = True Then RstCab("asigfam") = -1
    
    RstCab("basico") = NulosN(txt(12).Text)
    RstCab("bono") = NulosN(txt(13).Text)
    
    RstCab("paghornor") = NulosN(txt(15).Text)
    RstCab("paghorext") = NulosN(txt(16).Text)
    '----------------------------------------
    RstCab("idcat") = 0
    RstCab("fching") = Null
    RstCab("fchcese") = Null
    
    RstCab("nombre") = Trim(txt(1).Text) & " " & Trim(txt(2).Text) & " " & Trim(txt(3).Text)
    RstCab("nombre") = Replace(RstCab("nombre"), "  ", " ")
    
    RstCab.Update
    
    '--periodo laboral
    For A = 1 To Fg2.Rows - 1
        RstPerLab.AddNew
        RstPerLab("idemp") = xId
        RstPerLab("corr") = A
        RstPerLab("idcat") = Fg2.Cell(flexcpText, A, 2)
        If NulosN(Fg2.Cell(flexcpText, A, 2)) = 3 Then '--categoria -modalidad formativa
            RstPerLab("idmodfor") = NulosN(Fg2.TextMatrix(A, 3))   '--tipo convenio - modalidad formatva
        Else
            RstPerLab("idfinper") = NulosN(Fg2.Cell(flexcpText, A, 6)) '--tipo de extincion del contrato
        End If
        If IsDate(Fg2.TextMatrix(A, 4)) = True Then
            RstPerLab("fchini") = CDate(Fg2.TextMatrix(A, 4))
        End If
        If IsDate(Fg2.Cell(flexcpText, A, 5)) = True Then
            RstPerLab("fchfin") = CDate(Fg2.Cell(flexcpText, A, 5))
        End If
        RstPerLab.Update
    Next A

    '--insertando el centro de costo
    If Fg3.Rows > 1 Then
        For A = 1 To Fg3.Rows - 1
            RstCenCos.AddNew
            RstCenCos("idemp") = xId
            RstCenCos("idcencos") = NulosN(Fg3.TextMatrix(A, 4))
            RstCenCos("imppor") = NulosN(Fg3.TextMatrix(A, 3))
            RstCenCos.Update
        Next A
    End If
    '--actualizar la categoria, fecha de ingreso, fecha de cese
    nSQL = "UPDATE (SELECT pla_empleados.id, pla_empleados.idcat, pla_empleados.fching, pla_empleados.fchcese FROM pla_empleados WHERE (((pla_empleados.id)=" & xId & "))) AS emp " _
        + vbCr + " INNER JOIN (SELECT TOP 1 pla_periodolaboral.idemp, pla_periodolaboral.idcat, pla_periodolaboral.fchini, pla_periodolaboral.fchfin FROM pla_periodolaboral WHERE (((pla_periodolaboral.idemp)=" & xId & "))  ORDER BY pla_periodolaboral.fchini DESC ) AS periodo " _
        + vbCr + " ON emp.id = periodo.idemp " _
        + vbCr + " SET emp.idcat = [periodo].[idcat], emp.fching = [periodo].[fchini], emp.fchcese = [periodo].[fchfin];"

    XCON_.Execute nSQL
    '----------------------------------------------------------------
    'grabamos la foto, firma y la copiamos a su ruta final
    'Dim Rst As New ADODB.Recordset
    Dim xArchivo, xPath, xRuta As String
    'Dim xFile As New Scripting.FileSystemObject
    Dim xfile As Object
    Set xfile = CreateObject("Scripting.FileSystemObject")
    
    RST_Busq Rst, "SELECT TOP 1 * FROM pla_empleadosimg ", XCON_
    
    
    xPath = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
    xPath = xPath + "personal\"
    
    If pic(0).Tag <> "" Then
        xArchivo = Format(xId, "0000") & "-01.jpj"
        xRuta = xPath & xArchivo
        If ArchivoExiste(xRuta) = True Then xfile.DeleteFile xRuta
        xfile.CopyFile pic(0).Tag, xRuta
        
        Rst.AddNew
        Rst("idemp") = xId
        Rst("corr") = 1
        Rst("tipo") = 1
        Rst("descripcion") = xArchivo
        Rst.Update
    End If
    If pic(1).Tag <> "" Then
        xArchivo = Format(xId, "0000") & "-02.jpj"
        xRuta = xPath & xArchivo
        If ArchivoExiste(xRuta) = True Then xfile.DeleteFile xRuta
        xfile.CopyFile pic(1).Tag, xRuta
        Rst.AddNew
        Rst("idemp") = xId
        Rst("corr") = 2
        Rst("tipo") = 2
        Rst("descripcion") = xArchivo
        Rst.Update
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, Quehace, xHorIni, Time, Date, XCON_, xId

    
    '----------------------------------------------------------------
    XCON_.CommitTrans
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstCenCos = Nothing
    
    If PRIMERO_ Then
        MsgBox "Los datos del Personal " + IIf(Quehace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo
    End If

    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstCenCos = Nothing
    XCON_.RollbackTrans
    MsgBox "No se pudo guardar al Personal por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub nuevo()
    Quehace = 1
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    txt_cb(2).Text = 193
    txt_cb_Validate 2, False
        
    TabOne2.TabEnabled(2) = False
    
    Fg1.SelectionMode = flexSelectionFree
    Fg3.SelectionMode = flexSelectionFree
    
    opt_esplanilla(0).Value = True
    
    Label5.Caption = "Agregando Personal"
    
    '--agregar código personal
    DefinirCodigo
    
    TabOne2.CurrTab = 0
    txt(1).SetFocus
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Personal"

    ActivaTool
    
    Bloquea True
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
    End If
    
    Fg1.SelectionMode = flexSelectionFree
    Fg3.SelectionMode = flexSelectionFree
    
    TabOne1.TabEnabled(0) = False
    
    '--verificar si tiene asignado codigo, caso contrario se procedera a definirlo
    If NulosC(LblCodigoEmp.Caption) = "" Then DefinirCodigo
    
    Quehace = 2
    xHorIni = Time
    Agregando = False
    If TabOne2.CurrTab <> 0 Then TabOne2.CurrTab = 0
    txt(1).SetFocus

End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Eliminar(ByRef XCON_ As ADODB.Connection, NUEVO_ As Boolean)
    Dim Rpta As Integer
    Dim RstBus As New ADODB.Recordset
    
    TabOne1.CurrTab = 0
    
    If RstFrm.RecordCount = 0 Then
        If NUEVO_ Then
            MsgBox "No hay Registros para Eliminar", vbExclamation, xTitulo
        End If
        
        Exit Sub
    End If
    
    RST_Busq RstBus, "SELECT TOP 1 * FROM pla_boleta WHERE idemp = " & NulosN(RstFrm("id")) & "", xCon
    If RstBus.RecordCount <> 0 Then
        If NUEVO_ Then
            MsgBox "No se puede eliminar el personal especificado, figura en procesos de planilla", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        
        Set RstBus = Nothing
        Exit Sub
    End If
    '--validando si personal se encuentra en pago de producción
    RST_Busq RstBus, "SELECT TOP 1 * FROM pro_pagos where idemp =" & RstFrm("id") & "", xCon
    If RstBus.RecordCount <> 0 Then
        If NUEVO_ Then
            MsgBox "No se puede eliminar el personal especificado, figura en Registro de Tareas de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        
        Set RstBus = Nothing
        Exit Sub
    End If
        
    Set RstBus = Nothing
        
    If NUEVO_ Then
        Rpta = MsgBox("Esta seguro de eliminar al personal seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbNo Then Exit Sub
    End If
    
    XCON_.Execute "DELETE FROM pla_derechohab WHERE idemp=" & RstFrm.Fields("id") & " ;"
    XCON_.Execute "DELETE * FROM pla_empleados WHERE id = " & RstFrm("id") & ""
    
    'Eliminar historial del registro
    XCON_.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo
    
    If NUEVO_ Then
        MsgBox "El personal se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then nuevo

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then
        Eliminar xCon, True
        
        '*********************************
        If SINCPLANILLA_ Then
            procesarRestoEmpresas 2
        End If
        '*********************************
        
        RstFrm.Requery
        Dg1.Refresh
    End If

    If Button.Index = 5 Then Cancelar

    If Button.Index = 6 Then
        If Grabar(xCon, True) = True Then
            
            '*********************************
            If SINCPLANILLA_ Then
                ' Se Graba las demas Empresas
                procesarRestoEmpresas 1
            End If
            '*********************************
            
            RstFrm.Requery

            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If

            Dg1.Refresh
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
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(6, 4) As String
    Dim nSQLWhere As String
    
    xCampos(0, 0) = "TipDoc":               xCampos(0, 1) = "docabrev":   xCampos(0, 2) = "700":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Numero":               xCampos(1, 1) = "numdoc":     xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombres":    xCampos(2, 2) = "3200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Sexo":                 xCampos(3, 1) = "sexo":       xCampos(3, 2) = "550":    xCampos(3, 3) = "C"
    xCampos(4, 0) = "Categoría":            xCampos(4, 1) = "categoria":  xCampos(4, 2) = "2000":    xCampos(4, 3) = "C"
    xCampos(5, 0) = "Estado":               xCampos(5, 1) = "estado":     xCampos(5, 2) = "700":    xCampos(5, 3) = "C"

    nSQL = "SELECT pla_empleados.*, mae_dociden.abrev AS docabrev, mae_dociden.descripcion AS desciden, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_sexo.abrev AS sexo, IIf([pla_empleados].[fchcese] Is Not Null,'De Baja','Activo') AS estado " _
        + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex; "

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Personal", "nombres", "nombres", Principio

    If xRs.State = 1 Then
        RstFrm.MoveFirst
        RstFrm.Find "id = " & xRs("id") & ""
    End If
    
    Set xRs = Nothing
End Sub

Sub Filtrar()
    TabOne1.CurrTab = 0
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 3) As String

    xCampos(0, 0) = "Ape. y Nombres":     xCampos(0, 1) = "apenom":      xCampos(0, 2) = "C"
    xCampos(1, 0) = "Cargo":              xCampos(1, 1) = "descargo":    xCampos(1, 2) = "C"
    xCampos(2, 0) = "Tipo":               xCampos(2, 1) = "destipser":   xCampos(2, 2) = "C"
    xCampos(3, 0) = "Basico":             xCampos(3, 1) = "basico":      xCampos(3, 2) = "N"

    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub

Private Sub cb_Click(Index As Integer)
    If Quehace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--DOCUMENTO DE IDENTIDAD
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Documento de Identidad"
            nSQL = "SELECT mae_dociden.id, mae_dociden.descripcion as nombre, mae_dociden.id AS cod " _
                + vbCr + " From mae_dociden " _
                + vbCr + " ORDER BY mae_dociden.descripcion;"
        
        Case 1 '--SEXO
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Sexo"

             nSQL = "SELECT mae_sexo.id, mae_sexo.descripcion as nombre , mae_sexo.id AS cod " _
                + vbCr + " From mae_sexo " _
                + vbCr + " ORDER BY mae_sexo.descripcion;"
        
        Case 2 '--NACIONALIDAD
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Nacionalidad"

             nSQL = "SELECT mae_nacionalidad.id, mae_nacionalidad.descripcion as nombre, mae_nacionalidad.id AS cod " _
                + vbCr + " From mae_nacionalidad " _
                + vbCr + " ORDER BY mae_nacionalidad.descripcion;"
        
        Case 3 '--DEPARTAMENTO
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Departamento"

             nSQL = "SELECT mae_departamento.id, mae_departamento.descripcion as nombre, mae_departamento.id AS cod " _
                + vbCr + " From mae_departamento " _
                + vbCr + " ORDER BY mae_departamento.descripcion;"
        
        Case 4 '--PROVINCIA
            If NulosN(txt_cb(3).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(3).SetFocus
                Exit Sub
            End If
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Provincia"

             nSQL = "SELECT mae_provincia.id, mae_provincia.descripcion AS nombre, mae_provincia.id AS cod " _
                + vbCr + " From mae_provincia " _
                + vbCr + " Where (((mae_provincia.iddepa) = " & NulosN(txt_cb(3).Text) & " )) " _
                + vbCr + " ORDER BY mae_provincia.descripcion; "

        Case 5 '--DISTRITO
            If NulosN(txt_cb(3).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(3).SetFocus
                Exit Sub
            End If
            If NulosN(txt_cb(4).Text) = 0 Then
                MsgBox "Falta especificar la Provincia", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                Exit Sub
            End If
        
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Distrito"

             nSQL = "SELECT mae_distrito.id, mae_distrito.descripcion AS nombre, mae_distrito.id AS cod " _
                + vbCr + " From mae_distrito " _
                + vbCr + " Where (((mae_distrito.idprov) = " & NulosN(txt_cb(4).Text) & ")) " _
                + vbCr + " ORDER BY mae_distrito.descripcion;"
        
        Case 6 '--DOMICILIADO
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Si es Domiciliado"

             nSQL = "SELECT mae_indicadom.id, mae_indicadom.descripcion AS nombre, mae_indicadom.id AS cod " _
                + vbCr + " From mae_indicadom " _
                + vbCr + " ORDER BY mae_indicadom.descripcion;"
        
        Case 7 '--TIPO DE VIA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Tipo de Vía"

             nSQL = "SELECT mae_tipovia.id, mae_tipovia.descripcion AS nombre, mae_tipovia.id AS cod " _
                + vbCr + " From mae_tipovia " _
                + vbCr + " ORDER BY mae_tipovia.descripcion;"
        
        Case 8 '--TIPO ZONA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Tipo de Zona"

             nSQL = "SELECT mae_tipozona.id, mae_tipozona.descripcion AS nombre, mae_tipozona.id AS cod " _
                + vbCr + " From mae_tipozona " _
                + vbCr + " ORDER BY mae_tipozona.descripcion;"

        Case 9 '--ESSALUD + Vida
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando ESSALUD + Vida"

             nSQL = "SELECT mae_indicaesalud.id, mae_indicaesalud.descripcion AS nombre, mae_indicaesalud.id AS cod " _
                + vbCr + " From mae_indicaesalud " _
                + vbCr + " ORDER BY mae_indicaesalud.descripcion DESC;"
    
        Case 10 '--cargo
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Cargo del Personal"
            
            nSQL = "SELECT mae_cargo.id, mae_cargo.descripcion AS nombre, mae_cargo.id AS cod " _
                + vbCr + " FROM mae_cargo;"
        Case 11 '--area
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Area del Personal"
            
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod  " _
                + vbCr + " FROM mae_area; "
        Case 12 '--lugar de trabajo
            ReDim xCampos(4, 3) As String
            xCampos(0, 0) = "Establecimiento":  xCampos(0, 1) = "centro":  xCampos(0, 2) = "1500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombre":           xCampos(1, 1) = "nombre":           xCampos(1, 2) = "2500":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Dirección":        xCampos(2, 1) = "direc":            xCampos(2, 2) = "3000":   xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":               xCampos(3, 1) = "id":               xCampos(3, 2) = "400":    xCampos(3, 3) = "N"
            nTitulo = "Buscando Lugar de Trabajo"
            
            nSQL = "SELECT mae_empresalugtra.id, mae_empresalugtra.descripcion AS nombre, mae_empresalugtra.id AS cod, mae_empresalugtra.direc, mae_tipoestablecimiento.descripcion AS centro " _
                + vbCr + " FROM mae_tipoestablecimiento RIGHT JOIN mae_empresalugtra ON mae_tipoestablecimiento.id = mae_empresalugtra.idtipest;"
                
        Case 13 '--boleta de pago
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":             xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Tipo de Boletas de Pago"
            
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod, pla_proceso.abrev " _
                + vbCr + " From pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1));"
                
        '********************************************************************
        Case 14 ' Tipo de Planilla
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Tipo de Planilla"
            
            nSQL = "SELECT pla_tipoplanilla.id, pla_tipoplanilla.descripcion AS nombre, pla_tipoplanilla.id As cod " _
                + vbCr + "FROM pla_tipoplanilla;"
        '********************************************************************
            
    End Select

    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).Tag = xRs.Fields(1) & "" '--NOMBRE
    
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 3 '--DEPARTAMENTO
                txt_cb(4).Text = ""
                txt_cb(5).Text = ""
            Case 4 '--PROVINCIA
                txt_cb(5).Text = ""
        End Select
    End If
    Select Case Index
        Case 0 '--DOC IDENTIDAD
            txt(4).SetFocus
        Case 1 '--SEXO
            txt_cb(2).SetFocus
        Case 2 '--NACIONALIDAD
            txt(5).SetFocus
        Case 3 '--DEPARTAMENTO
            txt_cb(4).SetFocus
        Case 4 '--PROVINCIA
            txt_cb(5).SetFocus
        Case 5 '--DISTRITO
            txt_cb(6).SetFocus
        Case 6 '--DOMICIALIDO
            txt_cb(7).SetFocus
        Case 7 '--TIPO VIA
            txt(7).SetFocus
        Case 8 '--TIPO ZONA
            txt(10).SetFocus
        Case 9 '--ESSALUD-VIDA
            txt(14).SetFocus
            
        Case 10 '--cargo
            txt_cb(12).SetFocus
        Case 11 '--area
            txt_cb(10).SetFocus
        Case 12 '--lugar de trabajo
            txt_cb(13).SetFocus
        Case 13 '--boleta de trabajo
    End Select

Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If Quehace = 3 Or Quehace = -1 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        lbl_cb(Index).Caption = ""
        lbl_cod(Index).Caption = ""
        lbl_cb(Index).Tag = ""
        If Index = 3 Then '--departamento
            txt_cb(4).Text = ""
            txt_cb(5).Text = ""
        ElseIf Index = 4 Then '--provincia
            txt_cb(5).Text = ""
        End If
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Quehace = 3 Then Exit Sub
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
    If Quehace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index

        Case 0 '--DOCUMENTO DE IDENTIDAD
            nSQL = "SELECT mae_dociden.id, mae_dociden.descripcion as nombre, mae_dociden.id AS cod " _
                + vbCr + " FROM mae_dociden " _
                + vbCr + " WHERE mae_dociden.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 1 '--SEXO
             nSQL = "SELECT mae_sexo.id, mae_sexo.descripcion as nombre , mae_sexo.id AS cod " _
                + vbCr + " From mae_sexo " _
                + vbCr + " WHERE mae_sexo.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 2 '--NACIONALIDAD
             nSQL = "SELECT mae_nacionalidad.id, mae_nacionalidad.descripcion as nombre, mae_nacionalidad.id AS cod " _
                + vbCr + " From mae_nacionalidad " _
                + vbCr + " WHERE mae_nacionalidad.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 3 '--DEPARTAMENTO
             nSQL = "SELECT mae_departamento.id, mae_departamento.descripcion as nombre, mae_departamento.id AS cod " _
                + vbCr + " From mae_departamento " _
                + vbCr + " WHERE mae_departamento.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 4 '--PROVINCIA
            If NulosN(txt_cb(3).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(3).SetFocus
                GoTo Salir
            End If
             nSQL = "SELECT mae_provincia.id, mae_provincia.descripcion AS nombre, mae_provincia.id AS cod " _
                + vbCr + " From mae_provincia " _
                + vbCr + " Where (((mae_provincia.iddepa) = " & NulosN(txt_cb(3).Text) & " )) " _
                + vbCr + " AND mae_provincia.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 5 '--DISTRITO
            If NulosN(txt_cb(3).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(3).SetFocus
                GoTo Salir
            End If
            If NulosN(txt_cb(4).Text) = 0 Then
                MsgBox "Falta especificar la Provincia", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                GoTo Salir
            End If

             nSQL = "SELECT mae_distrito.id, mae_distrito.descripcion AS nombre, mae_distrito.id AS cod " _
                + vbCr + " From mae_distrito " _
                + vbCr + " WHERE (((mae_distrito.idprov) = " & NulosN(txt_cb(4).Text) & "))  " _
                + vbCr + " AND mae_distrito.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 6 '--DOMICILIADO
             nSQL = "SELECT mae_indicadom.id, mae_indicadom.descripcion AS nombre, mae_indicadom.id AS cod " _
                + vbCr + " From mae_indicadom " _
                + vbCr + " WHERE mae_indicadom.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 7 '--TIPO DE VIA
             nSQL = "SELECT mae_tipovia.id, mae_tipovia.descripcion AS nombre, mae_tipovia.id AS cod " _
                + vbCr + " From mae_tipovia " _
                + vbCr + " WHERE mae_tipovia.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 8 '--TIPO ZONA
             nSQL = "SELECT mae_tipozona.id, mae_tipozona.descripcion AS nombre, mae_tipozona.id AS cod " _
                + vbCr + " From mae_tipozona " _
                + vbCr + " WHERE mae_tipozona.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 9 '--ESSALUD + Vida
             nSQL = "SELECT mae_indicaesalud.id, mae_indicaesalud.descripcion AS nombre, mae_indicaesalud.id AS cod " _
                + vbCr + " From mae_indicaesalud " _
                + vbCr + " WHERE mae_indicaesalud.id = " & NulosN(txt_cb(Index).Text) & ";"
        Case 10 '--CARGO
            nSQL = "SELECT mae_cargo.id, mae_cargo.descripcion AS nombre, mae_cargo.id AS cod " _
                + vbCr + " FROM mae_cargo " _
                + vbCr + " WHERE mae_cargo.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        Case 11 '--area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod  " _
                + vbCr + " FROM mae_area " _
                + vbCr + " WHERE mae_area.id = " & NulosN(txt_cb(Index).Text) & ";"
       
        Case 12 '--lugar de trabajo
            nSQL = "SELECT mae_empresalugtra.id, mae_empresalugtra.descripcion AS nombre, mae_empresalugtra.id AS cod, mae_empresalugtra.direc, mae_tipoestablecimiento.descripcion AS centro " _
                + vbCr + " FROM mae_tipoestablecimiento RIGHT JOIN mae_empresalugtra ON mae_tipoestablecimiento.id = mae_empresalugtra.idtipest " _
                + vbCr + " WHERE mae_tipoestablecimiento.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        Case 13 '--boleta de pago
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod, pla_proceso.abrev " _
                + vbCr + " From pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)) AND pla_proceso.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        '****************************************************
        Case 14 ' Tipo de Planilla
            nSQL = "SELECT pla_tipoplanilla.id, pla_tipoplanilla.descripcion AS nombre, pla_tipoplanilla.id As cod " _
                + vbCr + "FROM pla_tipoplanilla " _
                + vbCr + "WHERE (pla_tipoplanilla.id = " & NulosN(txt_cb(Index).Text) & ");"
        '****************************************************
    
    End Select

    If xCon.State = 0 Then GoTo Salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).Tag = RstTmp.Fields(1) & "" '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 3 '--DEPARTAMENTO
                txt_cb(4).Text = ""
                txt_cb(5).Text = ""
            Case 4 '--PROVINCIA
                txt_cb(5).Text = ""
            Case 1
'                txt(1).SetFocus
        End Select
    End If
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
Salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

Sub MuestraSegundoTab()
    Dim QueHaceTmp As Integer
    Blanquea
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    TabOne2.CurrTab = 0
    txt(0).Text = NulosN(RstFrm("id"))
    LblCodigoEmp.Caption = Format(NulosC(RstFrm("codigo")), "00000")
    
    txt(4).Text = NulosC(RstFrm("numdoc"))
    
    txt(1).Text = NulosC(RstFrm("apepat"))
    txt(2).Text = NulosC(RstFrm("apemat"))
    txt(3).Text = NulosC(RstFrm("nom"))
    txt(5).Text = NulosC(RstFrm("numtel"))
    txt(6).Text = NulosC(RstFrm("email"))
    txt(7).Text = NulosC(RstFrm("nomvia"))
    txt(8).Text = NulosC(RstFrm("numvia"))
    txt(9).Text = NulosC(RstFrm("intvia"))
    txt(10).Text = NulosC(RstFrm("nomzon"))
    txt(14).Text = NulosC(RstFrm("numessalud"))
    
    If NulosN(RstFrm("idtipdoc")) <> 0 Then
        txt_cb(0).Text = NulosN(RstFrm("idtipdoc"))
        lbl_cb(0).Caption = NulosC(RstFrm("desciden"))
        lbl_cod(0).Caption = NulosN(RstFrm("idtipdoc"))
    End If
    
    If NulosN(RstFrm("iddep")) <> 0 Then
        txt_cb(3).Text = NulosN(RstFrm("iddep"))
        lbl_cb(3).Caption = NulosC(RstFrm("depa"))
        lbl_cod(3).Caption = NulosN(RstFrm("iddep"))
    End If

    If NulosN(RstFrm("idpro")) <> 0 Then
        txt_cb(4).Text = NulosN(RstFrm("idpro"))
        lbl_cb(4).Caption = NulosC(RstFrm("prov"))
        lbl_cod(4).Caption = NulosN(RstFrm("idpro"))
    End If
    
    If NulosN(RstFrm("iddis")) <> 0 Then
        txt_cb(5).Text = NulosN(RstFrm("iddis"))
        lbl_cb(5).Caption = NulosC(RstFrm("dist"))
        lbl_cod(5).Caption = NulosN(RstFrm("iddis"))
    End If
    
    If NulosN(RstFrm("idarea")) <> 0 Then
        txt_cb(11).Text = NulosN(RstFrm("idarea"))
        lbl_cb(11).Caption = NulosC(RstFrm("area"))
        lbl_cod(11).Caption = NulosN(RstFrm("idarea"))
    End If
    
    If NulosN(RstFrm("idcargo")) <> 0 Then
        txt_cb(10).Text = NulosN(RstFrm("idcargo"))
        lbl_cb(10).Caption = NulosC(RstFrm("cargo"))
        lbl_cod(10).Caption = NulosN(RstFrm("idcargo"))
    End If
    
    txt(12).Text = Format(NulosN(RstFrm("basico")), FORMAT_MONTO)
    txt(13).Text = Format(NulosN(RstFrm("bono")), FORMAT_MONTO)
    
    txt(15).Text = Format(NulosN(RstFrm("paghornor")), FORMAT_MONTO)
    txt(16).Text = Format(NulosN(RstFrm("paghorext")), FORMAT_MONTO)
    
    '-----------------------
    QueHaceTmp = Quehace
    Quehace = -1 '--comodin para entrar a [txt_cb_Validate]
    
    If IsDate(RstFrm("fchnac") & "") = True Then
        txtfecha(0).Valor = CDate(RstFrm("fchnac"))
    End If
    If NulosN(RstFrm("idsex")) <> 0 Then
        txt_cb(1).Text = NulosN(RstFrm("idsex"))
        txt_cb_Validate 1, False
    End If
    If NulosN(RstFrm("idnac")) <> 0 Then
        txt_cb(2).Text = NulosN(RstFrm("idnac"))
        txt_cb_Validate 2, False
    End If
    If NulosN(RstFrm("indessalud")) <> 0 Then
        txt_cb(9).Text = NulosN(RstFrm("indessalud"))
        txt_cb_Validate 9, False
    End If
    If NulosN(RstFrm("inddomi")) <> 0 Then
        txt_cb(6).Text = NulosN(RstFrm("inddomi"))
        txt_cb_Validate 6, False
    End If
    If NulosN(RstFrm("idtipvia")) <> 0 Then
        txt_cb(7).Text = NulosN(RstFrm("idtipvia"))
        txt_cb_Validate 7, False
    End If
    If NulosN(RstFrm("idtipzon")) <> 0 Then
        txt_cb(8).Text = NulosN(RstFrm("idtipzon"))
        txt_cb_Validate 8, False
    End If
    
    
    If NulosN(RstFrm("idlugtra")) <> 0 Then
        txt_cb(12).Text = NulosN(RstFrm("idlugtra"))
        txt_cb_Validate 12, False
    End If
    
    If NulosN(RstFrm("idbolpag")) <> 0 Then
        txt_cb(13).Text = NulosN(RstFrm("idbolpag"))
        txt_cb_Validate 13, False
    End If
    
    '************************************************************
    If NulosN(RstFrm("idtippla")) <> 0 Then
        txt_cb(14).Text = NulosN(RstFrm("idtippla"))
        txt_cb_Validate 14, False
    End If
    '************************************************************
    
    If NulosN(RstFrm("aplanilla")) = -1 Then
        opt_esplanilla(1).Value = True
    Else
        opt_esplanilla(0).Value = True
    End If
    
    If NulosN(RstFrm("asigfam")) = -1 Then
        opt_asigfamiliar(1).Value = True
    Else
        opt_asigfamiliar(0).Value = True
    End If
    

    Quehace = QueHaceTmp
    
    lbl_persona(0).Caption = ">> " + StrConv(NulosC(RstFrm.Fields("nombres")), 3)
    lbl_persona(1).Caption = lbl_persona(0).Caption
    lbl_persona(2).Caption = lbl_persona(0).Caption
    
    pCargarDatosDerechoHabiente
    pCargarDatosPeriodoLaboral
    pCargarDatosCentroCosto
    '--colocando las fotos
    Dim xPath  As String
    Dim RstTmp As New ADODB.Recordset
    xPath = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
    xPath = xPath + "personal\"
    RST_Busq RstTmp, "select * from pla_empleadosimg where idemp = " & RstFrm("id"), xCon
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        RstTmp.Find "tipo=1"
        If RstTmp.EOF = False And RstTmp.BOF = False Then
            pic(0).Picture = LoadPicture(xPath & RstTmp("descripcion"))
        End If
    
        RstTmp.MoveFirst
        RstTmp.Find "tipo=2"
        If RstTmp.EOF = False And RstTmp.BOF = False Then
            pic(1).Picture = LoadPicture(xPath & RstTmp("descripcion"))
        End If
    End If
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String

    '******************************************************************************
    nSQL = "SELECT pla_empleados.*, mae_dociden.abrev AS docabrev, mae_dociden.descripcion AS desciden, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_sexo.abrev AS sexo, mae_cargo.descripcion AS cargo, mae_categoria.descripcion AS categoria, mae_categoria.nomcor AS catnomcorto, IIf([pla_empleados].[fchcese] Is Not Null,'De Baja','Activo') AS estado, mae_departamento.descripcion AS depa, mae_provincia.descripcion AS prov, mae_distrito.descripcion AS dist, mae_area.descripcion AS area, pla_empleados.fching & '' AS fching1, pla_empleados.fchcese & '' AS fchcese1, pla_tipoplanilla.descripcion AS destippla " _
        + vbCr + "FROM (mae_sexo RIGHT JOIN (((((((mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id) LEFT JOIN mae_categoria ON pla_empleados.idcat = mae_categoria.id) LEFT JOIN mae_departamento ON pla_empleados.iddep = mae_departamento.id) LEFT JOIN mae_provincia ON pla_empleados.idpro = mae_provincia.id) LEFT JOIN mae_distrito ON pla_empleados.iddis = mae_distrito.id) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) ON mae_sexo.id = pla_empleados.idsex) LEFT JOIN pla_tipoplanilla ON pla_empleados.idtippla = pla_tipoplanilla.id " _
        + vbCr + "ORDER BY pla_empleados.nombre;"
    '******************************************************************************

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
    band = Validar(txt)
    TabOne2.CurrTab = 0
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If
    
    band = Validar(txt_cb)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_capt(band).Caption, vbInformation, xTitulo
       txt_cb(band).SetFocus
       Exit Function
    End If

''    If IsDate(txtfecha(0).Valor) = False Then
''        MsgBox "Falta especificar la fecha de nacimiento" + vbCr + "Si no lo Conoce colocar [01/01/1900] por defecto", vbExclamation, xTitulo
''        txtfecha(0).SetFocus
''        Exit Function
''    End If
    
    If Trim(txt(6).Text) <> "" Then
        If InStr(txt(6).Text, "@") = 0 Or InStr(txt(6).Text, ".") = 0 Then
            MsgBox "El Campo e-mail es incorrecto", vbExclamation, xTitulo
            txt(6).SetFocus
            Exit Function
        End If
    End If
    '--DEL PERIODO LABORAL
''    If Fg2.Rows = 1 Then
''        MsgBox "Ingrese el Periodo Laboral", vbExclamation, xTitulo
''        TabOne2.CurrTab = 1
''        CmdPerLab(0).SetFocus
''        Exit Function
''    End If
    Dim mRow&, mCol&
    mCol = -1
    For mRow = 1 To Fg2.Rows - 1
        If NulosN(Fg2.Cell(flexcpText, mRow, 2)) = 0 Then '--categoria
            MsgBox "Falta especificar la Categoría", vbExclamation, xTitulo
            mCol = 2:          Exit For
        End If
        If NulosN(Fg2.Cell(flexcpText, mRow, 2)) = 3 And NulosN(Fg2.Cell(flexcpText, mRow, 3)) = 0 Then '--categoria - modalidad formativa
            MsgBox "Falta especificar el tipo de convenio de Modalidad Formativa", vbExclamation, xTitulo
            mCol = 3:          Exit For
        End If
        If IsDate(Fg2.TextMatrix(mRow, 4)) = False Then
            MsgBox "Falta especificar la Fecha de Inicio o Reinicio", vbExclamation, xTitulo
            mCol = 4:          Exit For
        End If
        If mRow < Fg2.Rows - 1 Then '--si la fila actual es inferior al total de filas => obligar que se ingrese los datos
            If IsDate(Fg2.TextMatrix(mRow, 5)) = False Then
                MsgBox "Falta especificar la Fecha de Fin, Cese / Suspensión", vbExclamation, xTitulo
                mCol = 5:          Exit For
            End If
            If NulosN(Fg2.Cell(flexcpText, mRow, 2)) <> 3 And NulosN(Fg2.Cell(flexcpText, mRow, 6)) = 0 Then '--categoria - modalidad formativa
                MsgBox "Falta especificar el tipo de Extinción del Contrato", vbExclamation, xTitulo
                mCol = 6:          Exit For
            End If
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg2.Row = mRow: Fg2.Col = mCol: Agregando = False
        TabOne2.CurrTab = 1
        Fg2.SetFocus
        Exit Function
    End If
   
    '--
    If opt_esplanilla(0).Value = False And opt_esplanilla(1).Value = False Then
        TabOne2.CurrTab = 3
        MsgBox "Seleccione si está en planilla", vbExclamation, xTitulo
        Exit Function
    End If
    
'    If opt_esplanilla(1).Value = True And NulosN(txt(12).Text) = 0 Then
'        TabOne2.CurrTab = 3
'        MsgBox "Ingrese el sueldo Básico", vbExclamation, xTitulo
'        txt(12).SetFocus
'        Exit Function
'    ElseIf opt_esplanilla(0).Value = True And NulosN(txt(13).Text) = 0 Then
'        TabOne2.CurrTab = 3
'        MsgBox "Ingrese la Bonificación", vbExclamation, xTitulo
'        txt(13).SetFocus
'        Exit Function
'    End If
    
    '*************************************************************************
    '--del centro de costo
    pTotalizarCenCos '--recalcular el porcentaje total
    If NulosN(txt_CenCos.Text) <> 0 And NulosN(txt_CenCos.Text) < 100 Then
        MsgBox "El Porcentaje Acumulado no llega al 100%" + vbCr + "Modifique los valores ingresados", vbExclamation, xTitulo
        Agregando = True: Fg3.Row = 1:  Fg3.Col = 3:    Agregando = False
        Fg3.SetFocus
        Exit Function
    End If
    mCol = -1
    For mRow = Fg3.FixedRows To Fg3.Rows - 1
        If NulosN(Fg3.TextMatrix(mRow, 4)) = 0 Then   '--codigo centro costo
            MsgBox "Falta especificar el Centro de Costo", vbExclamation, xTitulo
            mCol = 1:          Exit For
        ElseIf NulosN(Fg3.TextMatrix(mRow, 3)) = 0 Then '--porcentaje
            MsgBox "Falta especificar el Porcentaje", vbExclamation, xTitulo
            mCol = 3:          Exit For
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg3.Row = mRow: Fg3.Col = mCol: Agregando = False
        Fg3.SetFocus
        Exit Function
    End If

    '*************************************************************************
    
    '--
    fValidarDatos = True
    
End Function

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 4 '--
            Select Case NulosN(txt_cb(0).Text)
                Case 1, 5, 15, 16 '--DNI,RUC
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Case Else
                    
            End Select
        Case 5 '--telef
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
            
    End Select
End Sub


Public Sub pCargarDatosDerechoHabiente()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    Fg1.Rows = 1
    '--verificando que existan parientes
    nSQL = "SELECT * FROM pla_derechohab Where pla_derechohab.idemp = " & RstFrm.Fields("id")
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount = 0 Then
        Set RstTmp = Nothing
        Exit Sub
    End If
    Set RstTmp = Nothing
    '-----
    nSQL = "SELECT pla_derechohab.corr, pla_derechohab.apepat & ' ' & pla_derechohab.apemat & ' ' & pla_derechohab.nombre AS nombres, mae_sexo.abrev AS sexo, mae_vinculofam.descripcion AS vinculo, mae_dociden.abrev AS docabrev, pla_derechohab.numdoc, pla_derechohab.fchnac " _
        + vbCr + " FROM mae_vinculofam RIGHT JOIN (mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_derechohab ON mae_dociden.id = pla_derechohab.idtipdoc) ON mae_sexo.id = pla_derechohab.idsex) ON mae_vinculofam.id = pla_derechohab.idvinfam " _
        + vbCr + " Where (((pla_derechohab.idemp) = " & RstFrm.Fields("id") & ")) " _
        + vbCr + " ORDER BY pla_derechohab.apepat & ' ' & pla_derechohab.apemat & ' ' & pla_derechohab.nombre;"
 
    RST_Busq RstTmp, nSQL, xCon
    
    Agregando = True
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("corr"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("vinculo"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstTmp("nombres"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstTmp("sexo"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp("fchnac"), "dd/mm/yy")
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(RstTmp("docabrev"))
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(RstTmp("numdoc"))
        RstTmp.MoveNext
    Loop
    Agregando = False
    Set RstTmp = Nothing
End Sub


'**************************
'**** PERIODO LABORAL
'**************************

Private Sub pConfigurarGrilla()
    With Fg1 '--DERECHOHABIENTE
        .Rows = 1
        .Cols = 8
        .FixedRows = 1
        .RowHeight(0) = 250
        .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Vínculo":          .ColWidth(2) = 1100:     .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Nombres":          .ColWidth(3) = 4500:    .ColAlignment(3) = flexAlignLeftCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Sexo":             .ColWidth(4) = 540:     .ColAlignment(4) = flexAlignCenterCenter:   .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Fch. Nac.":        .ColWidth(5) = 1245:    .ColAlignment(5) = flexAlignCenterCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Tipo Doc.":        .ColWidth(6) = 1200:    .ColAlignment(6) = flexAlignCenterCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 7) = "N°.Documento":     .ColWidth(7) = 2000:    .ColAlignment(7) = flexAlignLeftCenter:   .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
        .ColFormat(4) = "##/##/####"
        .SelectionMode = flexSelectionByRow
    End With
    With Fg2 '--PERIODO LABORAL
        .Rows = 1
        .Cols = 9
        .ColWidth(1) = 200
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Categoría":                                    .ColWidth(2) = 2500:    .ColAlignment(2) = flexAlignLeftCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Tipo Convenio" + vbCr + "(Solo Modalidad Formativa)":    .ColWidth(3) = 2500:     .ColAlignment(3) = flexAlignLeftCenter:         .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Fch.Inicio" + vbCr + "o Reinicio":        .ColWidth(4) = 1100:     .ColAlignment(4) = flexAlignCenterCenter:       .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Fch.Fin, Cese" + vbCr + "/ Suspensión":    .ColWidth(5) = 1100:    .ColAlignment(5) = flexAlignLeftCenter:         .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Tipo de Extinción del Contrato" + vbCr + "(No Considerar Modalidad Formativa)":   .ColWidth(6) = 2700:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 7) = "...":   .ColWidth(7) = 650:    .ColAlignment(7) = flexAlignLeftCenter:         .Row = 0: .Col = 7: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 8) = "Estado":   .ColWidth(8) = 650:    .ColAlignment(8) = flexAlignLeftCenter:         .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
        .ColEditMask(4) = "##/##/####"
        .ColEditMask(5) = "##/##/####"
        .SelectionMode = flexSelectionFree
    End With
    '*****************************************
    Fg3.SelectionMode = flexSelectionByRow
    Fg3.ColWidth(4) = 0
    OCULTAR_COL Fg3, 4, 4
    GRID_COMBOLIST Fg3, 1
    '*****************************************
    GRID_COMBOLIST Fg2, 4
    GRID_COMBOLIST Fg2, 5
    GRID_COMBOLIST Fg2, 7
    'Fg2.ColDataType(7) = flexDTBoolean
    '--COMBOLIST CON VSFLEXGRID
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
    '--categoria
    RST_Busq RstTmp, "SELECT mae_categoria.id, mae_categoria.descripcion FROM mae_categoria  ;", xCon
    tFormat = Fg2.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg2.ColComboList(2) = tFormat
    Set RstTmp = Nothing
    '--tipo convenio
    RST_Busq RstTmp, "SELECT mae_tipomodformativa.id, mae_tipomodformativa.descripcion FROM mae_tipomodformativa ORDER BY mae_tipomodformativa.descripcion;", xCon
    tFormat = Fg2.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg2.ColComboList(3) = tFormat
    Set RstTmp = Nothing
    '--tipo extincion del contrato
    RST_Busq RstTmp, "SELECT mae_finperiodo.id, mae_finperiodo.descripcion FROM mae_finperiodo ORDER BY mae_finperiodo.descripcion;", xCon
    tFormat = Fg2.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg2.ColComboList(6) = tFormat
    Set RstTmp = Nothing
        
    DoEvents
End Sub

Private Sub fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Quehace <> 3 Then
        If Col = 4 Or Col = 5 Then
            '--invocar al formulario de fecha
            Dim obj As New SGI2_funciones.formularios
            obj.FechaSeleccionar Fg2, Row, Col, Fg2.TextMatrix(Row, Col)
            Set obj = Nothing
        End If
    End If
    If Quehace <> 3 Then Exit Sub
    If Col <> 7 Then Exit Sub
    Select Case NulosN(Fg2.Cell(flexcpText, Row, 2))
        Case 1 '--trabajador
            FrmCatTrabajador.pRecibeLink 1
            FrmCatTrabajador.Show 1
        Case 2 '--pensionista
            FrmCatPensionista.pRecibeLink 1
            FrmCatPensionista.Show 1
        Case 3 '--prestador de servicio modalidad  formativa
            FrmCatModFormativa.pRecibeLink 1
            FrmCatModFormativa.Show 1
        Case 4 '--prestador de servicio 4ta cat
            FrmCat4taCategoria.Show 1
        Case 5 '--personal de tercero
            FrmCatPersonalTerceros.pRecibeLink 1
            FrmCatPersonalTerceros.Show 1
        Case Else
            
    End Select
End Sub

Private Sub fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    Select Case Col
        Case 2 '--categoria
            If Fg2.Cell(flexcpText, Row, Col) = "" Then
                Fg2.TextMatrix(Row, 3) = ""
                Fg2.TextMatrix(Row, 6) = ""
                Exit Sub
            End If
            If Fg2.Cell(flexcpText, Row, Col) = 3 Then '--categoria
                Fg2.TextMatrix(Row, 6) = ""
            Else
                Fg2.TextMatrix(Row, 3) = ""
            End If
        Case 3 '--tipo convenio
            If Fg2.Cell(flexcpText, Row, 2) = "" Then
                MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, 3) = ""
                Fg2.TextMatrix(Row, 6) = ""
                Fg2.Col = 2
                Fg2.SetFocus
                Exit Sub
            End If
            If Fg2.Cell(flexcpText, Row, 2) <> 3 Then '--categoria modalidad formativa
                Fg2.TextMatrix(Row, 3) = ""
            End If
            
        Case 4 '--fecha inicio
            If Fg2.TextMatrix(Row, Col) = "  /  /    " Or Fg2.TextMatrix(Row, Col) = "" Then Exit Sub
            If IsDate(Fg2.TextMatrix(Row, Col)) = False Then
                MsgBox "La Fecha de Inicio o Reinicio es incorrecta", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "dd/mm/yyyy")
        Case 5 '--fecha fin
            If Fg2.TextMatrix(Row, Col) = "  /  /    " Or Fg2.TextMatrix(Row, Col) = "" Then
                Fg2.TextMatrix(Row, 6) = ""
                Fg2.TextMatrix(Row, 8) = "Activo"
                Exit Sub
            End If
            If IsDate(Fg2.TextMatrix(Row, Col)) = False Then
                MsgBox "La Fecha de Fin, Cese / Suspención  es incorrecta", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
            
            If IsDate(Fg2.TextMatrix(Row, 4)) = True Then
                If CDate(Fg2.TextMatrix(Row, 4)) > CDate(Fg2.TextMatrix(Row, 5)) Then
                    MsgBox "La Fecha de Fin es inferior a la Fecha de Inicio", vbExclamation, xTitulo
                    Fg2.TextMatrix(Row, Col) = ""
                    Exit Sub
                End If
            End If
            
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "dd/mm/yyyy")
            Fg2.TextMatrix(Row, 8) = "Baja"
        Case 6 '--tipo de extincion del contrato
            If Fg2.Cell(flexcpText, Row, 2) = "" Then
                MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, 3) = ""
                Fg2.TextMatrix(Row, 6) = ""
                Fg2.Col = 2
                Fg2.SetFocus
                Exit Sub
            End If
            If Fg2.Cell(flexcpText, Row, 2) = 3 Then '--categoria modalidad formativa
                Fg2.TextMatrix(Row, 6) = ""
            End If
    
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg2_CellChanged"
End Sub

Private Sub fg2_EnterCell()
    If Fg2.Row < 1 Then Exit Sub
    If Fg2.Col <> 7 And Quehace = 3 Then
        Fg2.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg2.Row <> Fg2.Rows - 1 Then
        Fg2.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg2.Col = 3 Then
        If Fg2.Cell(flexcpText, Fg2.Row, 2) = 3 Then  '--categoria modalidad formativa
            Fg2.Editable = flexEDKbdMouse
        Else
            Fg2.Editable = flexEDNone
        End If
    ElseIf Fg2.Col = 6 Then
        If NulosN(Fg2.Cell(flexcpText, Fg2.Row, 2)) = 3 Or NulosN(Fg2.Cell(flexcpText, Fg2.Row, 2)) = 0 Or IsDate(Fg2.TextMatrix(Fg2.Row, 5)) = False Then '--categoria modalidad formativa
            Fg2.Editable = flexEDNone
        Else
            Fg2.Editable = flexEDKbdMouse
        End If
    ElseIf Fg2.Col = 7 Then
        If Quehace = 3 Then
            Fg2.Editable = flexEDKbdMouse
        Else
            Fg2.Editable = flexEDNone
        End If
    Else
        Fg2.Editable = flexEDKbdMouse
    End If
    
End Sub

Private Sub fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Quehace = 3 Then Exit Sub
    If KeyCode = 45 Then
        pRegistroAdd
    End If
    If KeyCode = 46 Then
        pRegistroDel
    End If
End Sub

Public Sub pCargarDatosPeriodoLaboral()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error

    Agregando = True

    '************************************************************
       
    nSQL = "SELECT pla_periodolaboral.* From pla_periodolaboral Where (((pla_periodolaboral.idemp) = " & RstFrm.Fields("id") & ")) ORDER BY pla_periodolaboral.fchini ASC ; "

    RST_Busq RstTmp, nSQL, xCon
    
    Fg2.Rows = 1
    If RstTmp.RecordCount <> 0 Then
        CmdInfCat.Enabled = True
        RstTmp.MoveFirst
    Else
        CmdInfCat.Enabled = False
    End If
    Do While Not RstTmp.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosN(RstTmp("corr"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstTmp(("idcat")))
        If NulosC(RstTmp("idmodfor")) <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(RstTmp("idmodfor"))
        End If
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(RstTmp("fchini"))
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosC(RstTmp("fchfin"))
        If NulosC(RstTmp("idfinper")) <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = NulosC(RstTmp("idfinper"))
        End If
        '**********************************************************************************************
        Dim RstTmp1 As New ADODB.Recordset
        Select Case NulosC(RstTmp(("idcat")))
            Case 1 '--trabajador
                nSQL = "Select * from pla_categoria1 where idemp = " & RstFrm.Fields("id") & " ;"
            Case 2 '--pensionista
                nSQL = "Select * from pla_categoria2 where idemp = " & RstFrm.Fields("id") & " ;"
            Case 3 '--prestador de servicio - modalidad formativa
                nSQL = "Select * from pla_categoria4 where idemp = " & RstFrm.Fields("id") & " ;"
            Case 4 '--prestador de servicio - 4ta categoria
                nSQL = "Select * from pla_categoria3 where idemp = " & RstFrm.Fields("id") & " ;"
            Case 5 '--personal de tercero
                nSQL = "Select * from pla_categoria5 where idemp = " & RstFrm.Fields("id") & " ;"
            Case Else
                nSQL = ""
        End Select
        If nSQL = "" Then
            Fg2.TextMatrix(Fg2.Rows - 1, 7) = "Err"
        Else
            RST_Busq RstTmp1, nSQL, xCon
            If RstTmp1.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Rows - 1, 7) = "Ok"
            Else
                Fg2.TextMatrix(Fg2.Rows - 1, 7) = "Falta"
            End If
        End If
        If IsDate(RstTmp("fchfin")) = False Then
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = "Activo"
        Else
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = "Baja"
        End If
        
        Set RstTmp1 = Nothing
        '**********************************************************************************************
        RstTmp.MoveNext
    Loop
    Agregando = False
    Set RstTmp = Nothing
    Exit Sub
error:
    Agregando = False
    Set RstTmp = Nothing
    habilitar CmdPerLab, False
    CmdPerLab(3).Enabled = True
    SHOW_ERROR Me.Name, "pPonerDatos"
End Sub

Private Sub pRegistroAdd()
    Dim mCol%
    If Quehace = 3 Then Exit Sub
    Agregando = True
    If Fg2.Rows > 1 Then
        If NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 2)) = 0 Then  '--categoria
            MsgBox "Falta ingresar la Categoría", vbExclamation, xTitulo
            mCol = 2
        Else
            If NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 2)) = 3 And NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 3)) = 0 Then    '--categoria modalidad formativa
                    MsgBox "Falta ingresar el Tipo de Convenio Modalidad Formativa", vbExclamation, xTitulo
                    mCol = 3
            End If
        End If
        If mCol = 0 Then
            If IsDate(Fg2.TextMatrix(Fg2.Rows - 1, 4)) = False Then
                MsgBox "Falta ingresar la Fecha de Inicio", vbExclamation, xTitulo
                mCol = 4
            ElseIf IsDate(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = False Then
                MsgBox "Falta ingresar la Fecha de Cese", vbExclamation, xTitulo
                mCol = 5
            ElseIf NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 2)) <> 3 And NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 6)) = 0 Then
                    MsgBox "Falta ingresar el Tipo de Extinción del Contrato", vbExclamation, xTitulo
                    mCol = 6
            Else

                Fg2.AddItem ""
                mCol = 2
            End If
        End If
    Else
        Fg2.AddItem ""
        mCol = 2
    End If
    Fg2.Row = Fg2.Rows - 1
    Fg2.Col = mCol
    Fg2.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If Fg2.Rows = 1 Then Exit Sub
    If Fg2.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg2.RemoveItem Fg2.Row
    
End Sub

Private Sub CmdPerLab_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR
            pRegistroAdd
        Case 1 '--ELIMINAR
            pRegistroDel
    End Select
End Sub

Private Sub opt_esplanilla_Click(Index As Integer)
    If Index = 0 Then '--no
        txt(12).Enabled = False
        txt(12).Text = "0.00"
    Else
        txt(12).Enabled = True
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 12, 13, 15, 16
            txt(Index).Text = Format(NulosN(txt(Index).Text), FORMAT_MONTO)
    End Select
End Sub


'***********************************************************************************
'-- centro de costo

Private Sub CmdCenCos_Click(Index As Integer)

    If Index = 0 Then

        Dim Rst As New ADODB.Recordset
        Dim A, B As Integer
        Dim Encontro As Boolean
        Dim xFrm As New SGI2_funciones.formularios
        Set Rst = xFrm.SeleCentroCosto(xCon)

        'Set Rst = SeleCentroCosto(xCon)
        If Rst.State = 1 Then
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                Encontro = False
                For A = 1 To Rst.RecordCount
                    For B = 1 To Fg3.Rows - 1
                        If Fg3.TextMatrix(B, 4) = NulosN(Rst("idcencos")) Then
                            Encontro = True
                        End If
                    Next B

                    If Encontro = False Then
                        If Fg3.TextMatrix(Fg3.Rows - 1, 1) <> "" Then
                            Fg3.Rows = Fg3.Rows + 1
                        End If
                        Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosC(Rst("codigo"))
                        Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(Rst("descripcion"))
                        Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosN(Rst("idcencos"))
                    End If

                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                Next A
            End If
        End If
        Set xFrm = Nothing
    Else
        If Fg3.Row < 1 Then Exit Sub
        If Fg3.Rows = 1 Then
            MsgBox "No hay Registros", vbInformation, xTitulo
            Exit Sub
        End If
        Fg3.RemoveItem Fg3.Row
        pTotalizarCenCos
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        CmdCenCos_Click 0
    End If
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 3 Then
        Fg3.TextMatrix(Row, Col) = Format(NulosN(Fg3.TextMatrix(Row, Col)), FORMAT_MONTO)
        If GRID_SUMAR_COL(Fg3, 3) > 100 Then
            MsgBox "El Total es superior al 100%", vbInformation, xTitulo
            Fg3.TextMatrix(Row, Col) = ""
        End If
        pTotalizarCenCos
    End If
End Sub

Private Sub pTotalizarCenCos()
    txt_CenCos.Text = Format(GRID_SUMAR_COL(Fg3, 3), FORMAT_MONTO)
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Col <> 3 Then
        KeyAscii = 0
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If

End Sub

Private Sub Fg3_EnterCell()
    Fg3.Editable = flexEDNone
    If Quehace = 3 Then Exit Sub
    If Fg3.Col = 1 Or Fg3.Col = 3 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If Quehace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdCenCos_Click 0
    End If
    If KeyCode = 46 Then
        If Fg3.Rows = 1 Then
            MsgBox "No se han especificado centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Fg3.RemoveItem Fg3.Row
        Fg3.Select Fg3.Rows - 1, 1, Fg3.Rows - 1, 1
    End If
End Sub



Private Sub pCargarDatosCentroCosto()

    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Me.MousePointer = vbHourglass
    nSQL = "SELECT con_centrocosto.*, pla_empleadoscos.imppor, pla_empleadoscos.idcencos " _
        + vbCr + " FROM con_centrocosto INNER JOIN pla_empleadoscos ON con_centrocosto.id = pla_empleadoscos.idcencos " _
        + vbCr + " WHERE (((pla_empleadoscos.IdEmp) = " & RstFrm.Fields("id") & ")) " _
        + vbCr + " ORDER BY con_centrocosto.codigo; "

    RST_Busq RstTmp, nSQL, xCon
    '**********************************************************************************************************
    Agregando = True
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg3.Rows = Fg3.Rows + 1
        '------
        Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosN(RstTmp("codigo"))
        Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(RstTmp("descripcion"))
        Fg3.TextMatrix(Fg3.Rows - 1, 3) = Format(NulosN(RstTmp("imppor")), FORMAT_MONTO)
        Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosN(RstTmp("idcencos"))

        RstTmp.MoveNext
    Loop
    Agregando = False
    '**********************************************************************************************************
    Set RstTmp = Nothing

    pTotalizarCenCos '--totalizar datos

    Me.MousePointer = vbDefault
    
End Sub


Private Sub pAddFoto(TIPO As Integer, oPic As PictureBox)
    Dim objCommDlg As Object
    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set objCommDlg = CreateObject("MSComDlg.CommonDialog")
    
    objCommDlg.Filter = "Archivos jpg (*.jpg)|*.jpg"
    objCommDlg.DialogTitle = "Abrir Archivo"
    objCommDlg.ShowOpen
    
    If Trim(objCommDlg.FileName) = "" Then
        oPic.Picture = LoadPicture("")
        oPic.Tag = ""
        Set objCommDlg = Nothing
        Exit Sub
    End If
    oPic.Picture = LoadPicture(objCommDlg.FileName, False)
    oPic.Tag = objCommDlg.FileName
    
    Set fs = Nothing
    Set objCommDlg = Nothing
    
End Sub


Private Sub DefinirCodigo()
    '===================================================================================================
    'Creado : 16/04/11 Por: Johan Castro
    'Propósito: Permitir definir el codigo del personal cuando se agregue un nuevo personal o cuando se
    '           modifique y este no tenga codigo asignado
    '
    'Entradas:  Ninguno
    '
    'Resultados: Código de empleado definido por el sistema
    '
    '===================================================================================================
    
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT pla_empleados.id, pla_empleados.codigo FROM pla_empleados WHERE (((pla_empleados.codigo) Is Not Null)) ORDER BY pla_empleados.codigo DESC "

    RST_Busq xRst, nSQL, xCon
    If xRst.State = 1 Then
        If xRst.RecordCount = 0 Then
            LblCodigoEmp.Caption = "00001"
        Else
            LblCodigoEmp.Caption = Format(NulosN(xRst("codigo") + 1), "00000")
        End If
    End If
    Set xRst = Nothing
    
    DoEvents
End Sub

Private Sub pExportar()
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(15, 3) As String
    
    TabOne1.CurrTab = 0
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":                   xCampos(0, 1) = "id":           xCampos(0, 2) = 2:      xCampos(0, 3) = "500"
    xCampos(1, 0) = "Código":               xCampos(1, 1) = "codigo":       xCampos(1, 2) = 0:      xCampos(1, 3) = "900"
    xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":       xCampos(2, 2) = 0:      xCampos(2, 3) = "2500"
    xCampos(3, 0) = "Sexo":                 xCampos(3, 1) = "sexo":         xCampos(3, 2) = 0:      xCampos(3, 3) = "500"
    xCampos(4, 0) = "T.D.":                 xCampos(4, 1) = "docabrev":     xCampos(4, 2) = 0:      xCampos(4, 3) = "500"
    xCampos(5, 0) = "Num. Doc":             xCampos(5, 1) = "numdoc":       xCampos(5, 2) = 0:      xCampos(5, 3) = "1200"
    xCampos(6, 0) = "Fch. Nac.":            xCampos(6, 1) = "fchnac":       xCampos(6, 2) = 1:      xCampos(6, 3) = "1100"
    xCampos(7, 0) = "Fch.Ingreso":          xCampos(7, 1) = "fching":       xCampos(7, 2) = 1:      xCampos(7, 3) = "1100"
    xCampos(8, 0) = "Categoría":            xCampos(8, 1) = "catnomcorto":  xCampos(8, 2) = 0:      xCampos(8, 3) = "500"
    xCampos(9, 0) = "Area":                 xCampos(9, 1) = "area":         xCampos(9, 2) = 0:      xCampos(9, 3) = "1300"
    xCampos(10, 0) = "Cargo":               xCampos(10, 1) = "cargo":       xCampos(10, 2) = 0:     xCampos(10, 3) = "1300"
    xCampos(11, 0) = "Pago H.N.":           xCampos(11, 1) = "paghornor":   xCampos(11, 2) = 2:     xCampos(11, 3) = "800"
    xCampos(12, 0) = "Pago H.E.":           xCampos(12, 1) = "paghorext":   xCampos(12, 2) = 2:     xCampos(12, 3) = "800"
    xCampos(13, 0) = "Estado":              xCampos(13, 1) = "estado":      xCampos(13, 2) = 0:     xCampos(13, 3) = "900"
    xCampos(14, 0) = "Fch. Cese":           xCampos(14, 1) = "fchcese":     xCampos(14, 2) = 1:     xCampos(14, 3) = "1100"
    '**********************************************************************
    xCampos(15, 0) = "Tip. Planilla":       xCampos(15, 1) = "destippla":   xCampos(15, 2) = 1:     xCampos(15, 3) = "1100"
    '**********************************************************************
    '***************************
    'Set RstTmp = RstFrm.Clone
    Set RstTmp = RstFrm
    '***************************
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "NOMINA DE PERSONAL", "", "", "NOMINA DE PERSONAL", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub

