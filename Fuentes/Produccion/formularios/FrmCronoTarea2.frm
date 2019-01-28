VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCronoTarea2 
   Caption         =   "Produccion - Programacion de Tareas"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   FillColor       =   &H80000002&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea2.frx":23EC
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
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
                  Text            =   "Recetas del producto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Productos "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   12
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
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11860
      _cx             =   20920
      _cy             =   12515
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
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
      FrontTabForeColor=   -2147483630
      Caption         =   "  Consulta  |   Detalle  "
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
         Height          =   6660
         Left            =   -12420
         TabIndex        =   5
         Top             =   390
         Width           =   11775
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6075
            Left            =   30
            TabIndex        =   6
            Top             =   540
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   10716
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Ini."
            Columns(1).DataField=   "fchini"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Fin."
            Columns(2).DataField=   "fchfin"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo Produccion"
            Columns(3).DataField=   "destippro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Programador"
            Columns(4).DataField=   "apenom"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1482"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2355"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3757"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3678"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=9208"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=9128"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Programacion de Tareas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   0
            TabIndex        =   7
            Top             =   30
            Width           =   11655
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6660
         Left            =   45
         TabIndex        =   3
         Top             =   390
         Width           =   11775
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   6180
            Left            =   10
            TabIndex        =   8
            Top             =   450
            Width           =   11745
            _cx             =   20717
            _cy             =   10901
            _ConvInfo       =   1
            Appearance      =   1
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
            BackColorSel    =   -2147483635
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCronoTarea2.frx":277E
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Programacion de Tareas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   11655
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   2850
      Left            =   90
      TabIndex        =   9
      Top             =   7590
      Width           =   11655
      _cx             =   20558
      _cy             =   5027
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorSel    =   -2147483635
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
      Rows            =   50
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCronoTarea2.frx":29CA
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "FrmCronoTarea2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean
Dim RstLis As New ADODB.Recordset
Dim fOrdenLista As Boolean 'especifica el orden de la lista de la consulta
Dim xHorIni As Date 'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer 'INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String

Dim indicador As Double


Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLis
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLis.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLis("id")), xCon
    End If
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If QueHace <> 2 Then Exit Sub
    If Agregando = True Then Exit Sub
    
    Dim h() As String
    Dim tiempo As Double
    Dim personas As Double
    
    Dim columna_duracion As Integer
    Dim columna_horIni As Integer
    Dim columna_horFin As Integer
    Dim columna_numPer As Integer
    Dim columna_cantidad As Integer
    
    columna_duracion = Fg1.Cols - 13
    columna_horIni = Fg1.Cols - 12
    columna_horFin = Fg1.Cols - 11
    columna_numPer = Fg1.Cols - 10
    columna_cantidad = Fg1.Cols - 9
    
    If Col = columna_duracion Or Col = columna_horFin Or Col = columna_numPer Then
        h = Split(Fg1.TextMatrix(Fg1.Row, columna_duracion), ":")
        tiempo = (60 * Val(h(0))) + Val(h(1))
        personas = CDbl(Fg1.TextMatrix(Fg1.Row, columna_numPer))
    
        indicador = personas * tiempo
    End If
    
    If Col = columna_horIni Then
        h = Split(Fg1.TextMatrix(Fg1.Row, columna_horIni), ":")
        tiempo = (60 * Val(h(0))) + Val(h(1))
        indicador = tiempo
    End If
    
    If Col = columna_cantidad Then
        tiempo = Fg1.TextMatrix(Fg1.Row, Col)
        indicador = tiempo
    End If
    
    If Col > columna_cantidad Then
        tiempo = Fg1.TextMatrix(Fg1.Row, Col)
        indicador = tiempo
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If QueHace <> 2 Then Exit Sub
    If Agregando = True Then Exit Sub
    Dim h() As String
    Dim horas As Integer
    Dim minutos As Integer
    
    Dim nuevopersonas As Double
    Dim nuevohoras As Double
    
    Dim columna_duracion As Integer
    Dim columna_horIni As Integer
    Dim columna_horFin As Integer
    Dim columna_numPer As Integer
    Dim columna_idpro As Integer
    Dim columna_cantidad As Integer
    
    columna_duracion = Fg1.Cols - 13
    columna_horIni = Fg1.Cols - 12
    columna_horFin = Fg1.Cols - 11
    columna_numPer = Fg1.Cols - 10
    columna_idpro = Fg1.Cols - 4
    columna_cantidad = Fg1.Cols - 9

    If Col = columna_duracion Then
        h = Split(Fg1.TextMatrix(Fg1.Row, columna_duracion), ":")
        horas = Val(h(0))
        If UBound(h) > 0 Then minutos = Val(h(1)) Else minutos = 0
        Fg1.TextMatrix(Fg1.Row, columna_duracion) = Format(horas, "00") & ":" & Format(minutos, "00")
        
        Fg1.TextMatrix(Fg1.Row, columna_horFin) = Format(CDate(Fg1.TextMatrix(Fg1.Row, columna_horIni)) + CDate(Fg1.TextMatrix(Fg1.Row, columna_duracion)), "HH:mm")
        
        nuevohoras = (60 * horas) + minutos
        Fg1.TextMatrix(Fg1.Row, columna_numPer) = Int(indicador / nuevohoras)
        If Fg1.TextMatrix(Fg1.Row, columna_numPer) = 0 Then Fg1.TextMatrix(Fg1.Row, columna_numPer) = 1
        Fg1.TextMatrix(Fg1.Row, columna_numPer) = Format(Fg1.TextMatrix(Fg1.Row, columna_numPer), "00")
        
        rellenarHoras Fg1.Row + 1, Fg1.TextMatrix(Fg1.Row, columna_idpro), Fg1.TextMatrix(Fg1.Row, columna_horFin)
    End If
    
    If Col = columna_horFin Then
        h = Split(Fg1.TextMatrix(Fg1.Row, columna_horFin), ":")
        horas = Val(h(0))
        If UBound(h) > 0 Then minutos = Val(h(1)) Else minutos = 0
        Fg1.TextMatrix(Fg1.Row, columna_horFin) = Format(horas, "00") & ":" & Format(minutos, "00")
        
        Fg1.TextMatrix(Fg1.Row, columna_duracion) = Format(CDate(Fg1.TextMatrix(Fg1.Row, columna_horFin)) - CDate(Fg1.TextMatrix(Fg1.Row, columna_horIni)), "HH:mm")
        
        h = Split(Fg1.TextMatrix(Fg1.Row, columna_duracion), ":")
        horas = Val(h(0))
        If UBound(h) > 0 Then minutos = Val(h(1)) Else minutos = 0
        
        nuevohoras = (60 * horas) + minutos
        Fg1.TextMatrix(Fg1.Row, columna_numPer) = Int(indicador / nuevohoras)
        If Fg1.TextMatrix(Fg1.Row, columna_numPer) = 0 Then Fg1.TextMatrix(Fg1.Row, columna_numPer) = 1
        Fg1.TextMatrix(Fg1.Row, columna_numPer) = Format(Fg1.TextMatrix(Fg1.Row, columna_numPer), "00")
        
        rellenarHoras Fg1.Row + 1, Fg1.TextMatrix(Fg1.Row, columna_idpro), Fg1.TextMatrix(Fg1.Row, columna_horFin)
    End If

    If Col = columna_numPer Then
        nuevopersonas = Fg1.TextMatrix(Fg1.Row, columna_numPer)
        Fg1.TextMatrix(Fg1.Row, columna_numPer) = Format(Fg1.TextMatrix(Fg1.Row, columna_numPer), "00")
        minutos = indicador / nuevopersonas
        horas = Int(minutos / 60)
        minutos = minutos Mod 60
        Fg1.TextMatrix(Fg1.Row, columna_duracion) = Format(horas, "00") & ":" & Format(minutos, "00")
        
        Fg1.TextMatrix(Fg1.Row, columna_horFin) = Format(CDate(Fg1.TextMatrix(Fg1.Row, columna_horIni)) + CDate(Fg1.TextMatrix(Fg1.Row, columna_duracion)), "HH:mm")
        rellenarHoras Fg1.Row + 1, Fg1.TextMatrix(Fg1.Row, columna_idpro), Fg1.TextMatrix(Fg1.Row, columna_horFin)
    End If
    
    If Col = columna_horIni Then
        MsgBox "La fecha de Inicio de la Tarea no se puede modificar, modifiquela en el cronograma de Produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        horas = Int(indicador / 60)
        minutos = indicador Mod 60
        
        Fg1.TextMatrix(Fg1.Row, columna_horIni) = Format(horas, "00") & ":" & Format(minutos, "00")
    End If
    
    If Col = columna_cantidad Then
        MsgBox "La cantidad del producto no se puede modificar, modifiquela en el cronograma de Produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        Fg1.TextMatrix(Fg1.Row, columna_cantidad) = Format(indicador, "0.00")
    End If
    
    If Col > columna_cantidad Then
        MsgBox "Estos parametros no son modificables, cambielos en la Receta del Producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        Fg1.TextMatrix(Fg1.Row, Col) = Format(indicador, "0.00")
    End If
End Sub

Private Sub rellenarHoras(FILA As Integer, id As Integer, horFinBase As String)
    Dim idAux As Integer
    Dim duracionTar As String
    Dim horIniTar As String
    Dim horFinTar As String
    
    Dim columna_idpro As Integer
    Dim columna_duracion As Integer
    Dim columna_horIni As Integer
    Dim columna_horFin As Integer
    Dim columna_numPer As Integer
    
    columna_duracion = Fg1.Cols - 13
    columna_horIni = Fg1.Cols - 12
    columna_horFin = Fg1.Cols - 11
    columna_idpro = Fg1.Cols - 4
    
    horIniTar = horFinBase
    With Fg1
        idAux = .TextMatrix(FILA, columna_idpro)
        While (id = idAux)
            duracionTar = .TextMatrix(FILA, columna_duracion)
            'se llena la hora de inicio y de fin de la tarea
            .TextMatrix(FILA, columna_horIni) = horIniTar
            horFinTar = NulosC(Format(CDate(duracionTar) + CDate(horIniTar), "HH:MM"))
            .TextMatrix(FILA, columna_horFin) = horFinTar
            horIniTar = horFinTar
            FILA = FILA + 1
            idAux = NulosN(.TextMatrix(FILA, columna_idpro))
        Wend
    End With
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
    
        SeEjecuto = True
            
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
    
        '--ocultar el boton a agregar
        Toolbar1.Buttons(1).Visible = False
        
        cSQL = "SELECT pro_cronograma.*, mae_tipoproducto.descripcion AS destippro, [pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ', ' & [pla_empleados]![nom] AS apenom " _
            + vbCr + "FROM (pla_empleados RIGHT JOIN (pro_cronograma LEFT JOIN pro_emp ON pro_cronograma.idsup = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) LEFT JOIN mae_tipoproducto ON pro_cronograma.idtippro = mae_tipoproducto.id " _
            + vbCr + "ORDER BY pro_cronograma.fchini DESC"
        
        RST_Busq RstLis, cSQL, xCon
        
        Set Dg1.DataSource = RstLis
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    iniciarCampos
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 8000
        Me.Width = 12000
    End If
End Sub

Private Sub iniciarCampos()
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    Fg1.Editable = flexEDNone
    Fg1.MergeCells = flexMergeSpill
    
    Frame1.BackColor = &H8000000F
    
    Me.Height = 8000
    Me.Width = 12000
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Label1.Caption = "Detalle de la Programacion de Tareas"
    Fg1.Editable = flexEDNone
    TabOne1.CurrTab = 0
End Sub

Sub Bloquea()

End Sub

Private Sub llenarTarea(Rst As ADODB.Recordset, ByRef Fgrid As VSFlexGrid, FILA As Integer, columna As Integer, ByRef horIniTar As String, duracionTar As String, ByRef cantidad As Double)
    Dim horFinTar As String
    
    With Fgrid
        .Select FILA, columna, FILA, columna
        .FillStyle = flexFillRepeat
        .CellBackColor = &H80000013  '&H00000000&
        'se llena el detalle de la tarea
        .TextMatrix(FILA, columna) = NulosC(Rst("destar"))
        
        .Select FILA, columna + 1, FILA, .Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &H80FFFF     '&H00000000&
        'Se llena la duracion de la tarea
        .TextMatrix(FILA, columna + 1) = duracionTar
        'se llena la hora de inicio y de fin de la tarea
        .TextMatrix(FILA, columna + 2) = horIniTar
        horFinTar = NulosC(Format(CDate(duracionTar) + CDate(horIniTar), "HH:MM"))
        .TextMatrix(FILA, columna + 3) = horFinTar
        horIniTar = horFinTar
        'se llena el numero de personas de la tarea
        .TextMatrix(FILA, columna + 4) = Format(Rst("numper"), "00")
        'se llena la cantidad
        .TextMatrix(FILA, columna + 5) = Format(cantidad, "0.00")
        'Nose
        .TextMatrix(FILA, columna + 6) = Format(Rst("aplpor"), "0.00")
        'Otro nose
        If NulosN(Rst("aplpor")) <> 0 Then
            .TextMatrix(FILA, columna + 7) = Format((cantidad * ((Rst("aplpor") / 100))), "0.00")
        Else
            .TextMatrix(FILA, columna + 7) = Format(cantidad, "0.00")
        End If
        cantidad = NulosN(.TextMatrix(FILA, columna + 7))
        'Se llena el detalle del producto
        .TextMatrix(FILA, columna + 8) = NulosC(Rst("fchpro"))
        .TextMatrix(FILA, columna + 9) = NulosN(Rst("iditem"))
        .TextMatrix(FILA, columna + 10) = NulosN(Rst("idpro"))
        .TextMatrix(FILA, columna + 11) = NulosN(Rst("idtar"))
        .TextMatrix(FILA, columna + 12) = NulosN(Rst("factor"))
        .TextMatrix(FILA, columna + 13) = NulosN(Rst("orden"))
        
        .Rows = Fg1.Rows + 1
    End With
End Sub

Private Sub ConfigurarGrid(es_matpri As Boolean)
    If es_matpri Then
        Fg1.Cols = 18
        Fg1.ColWidth(0) = 0
        Fg1.RowHeight(0) = 300
        
        Fg1.TextMatrix(0, 1) = "Fch. Prod."
        Fg1.ColWidth(1) = 1000
        Fg1.TextMatrix(0, 2) = "Materia Prima"
        Fg1.ColWidth(2) = 1100
        Fg1.TextMatrix(0, 3) = "Producto"
        Fg1.ColWidth(3) = 800
        Fg1.TextMatrix(0, 4) = "Tarea"
        Fg1.ColWidth(4) = 3500
        Fg1.TextMatrix(0, 5) = "Duracion"
        Fg1.ColWidth(5) = 800
        Fg1.TextMatrix(0, 6) = "Hor.Ini"
        Fg1.ColWidth(6) = 600
        Fg1.TextMatrix(0, 7) = "Hor.Fin"
        Fg1.ColWidth(7) = 600
        Fg1.TextMatrix(0, 8) = "NºPer."
        Fg1.ColWidth(8) = 600
        Fg1.TextMatrix(0, 9) = "Cantidad"
        Fg1.ColWidth(9) = 800
        Fg1.TextMatrix(0, 10) = "%Redto"
        Fg1.ColWidth(10) = 700
        Fg1.TextMatrix(0, 11) = "Tot.Proces"
        Fg1.ColWidth(11) = 900
        
        Fg1.ColWidth(12) = 0
        Fg1.ColWidth(13) = 0
        Fg1.ColWidth(14) = 0
        Fg1.ColWidth(15) = 0
        Fg1.ColWidth(16) = 0
        Fg1.ColWidth(17) = 0
    Else
        Fg1.Cols = 17
        'Fg1.FrozenCols = 3
        Fg1.ColWidth(0) = 0
        Fg1.RowHeight(0) = 300
        
        Fg1.TextMatrix(0, 1) = "Fch. Prod."
        Fg1.ColWidth(1) = 900
        Fg1.TextMatrix(0, 2) = "Producto"
        Fg1.ColWidth(2) = 2000
        Fg1.TextMatrix(0, 3) = "Tarea"
        Fg1.ColWidth(3) = 3500
        Fg1.TextMatrix(0, 4) = "Duracion"
        Fg1.ColWidth(4) = 800
        Fg1.TextMatrix(0, 5) = "Hor.Ini"
        Fg1.ColWidth(5) = 600
        Fg1.TextMatrix(0, 6) = "Hor.Fin"
        Fg1.ColWidth(6) = 600
        Fg1.TextMatrix(0, 7) = "NºPer."
        Fg1.ColWidth(7) = 600
        Fg1.TextMatrix(0, 8) = "Cantidad"
        Fg1.ColWidth(8) = 800
        Fg1.TextMatrix(0, 9) = "%Redto"
        Fg1.ColWidth(9) = 700
        Fg1.TextMatrix(0, 10) = "Tot.Proces"
        Fg1.ColWidth(10) = 900
        
        Fg1.ColWidth(11) = 0
        Fg1.ColWidth(12) = 0
        Fg1.ColWidth(13) = 0
        Fg1.ColWidth(14) = 0
        Fg1.ColWidth(15) = 0
        Fg1.ColWidth(16) = 0
    End If
End Sub

Sub MostrarSegundoTab()
    Dim RstDet As New ADODB.Recordset
    Dim Rpta As Integer
    
    Dim xTiempo As String
    'declaracion de las filas
    Dim f_fchini As Integer
    Dim f_matpri As Integer
    Dim f_despro As Integer
    Dim f_destar As Integer
    'declaracion de los valores auxiliares
    Dim fchiniAux As String
    Dim matpriAux As String
    Dim iditemAux As Integer
    Dim desproAux As String
    Dim idproAux As Integer
    Dim horiniAux As String
    Dim canproAux As Double
    
    
    Dim horprodAux As String
    
    Dim es_matpri As Boolean
    
    'se verifica si es cronograma de materia prima o no
    If RstLis("idtippro") = 1 Then
        es_matpri = True
    Else
        es_matpri = False
    End If
    'se configura la grilla segun si es materia prima o no
    ConfigurarGrid es_matpri
    
    Fg1.Rows = 2
    
    cSQL = "SELECT pro_cronogramatarea.*, alm_inventario.descripcion AS matpri, alm_inventario_1.descripcion AS despro, pro_tareas.descripcion AS destar " _
        + vbCr + "FROM ((pro_cronogramatarea LEFT JOIN alm_inventario ON pro_cronogramatarea.iditem = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_cronogramatarea.idpro = alm_inventario_1.id) LEFT JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id " _
        + vbCr + "Where (((pro_cronogramatarea.id) = " & RstLis("id") & ")) " _
        + vbCr + "ORDER BY pro_cronogramatarea.fchpro, alm_inventario.descripcion, alm_inventario_1.descripcion, pro_cronogramatarea.orden"

    RST_Busq RstDet, cSQL, xCon

    If RstDet.RecordCount = 0 Then
        Rpta = MsgBox("No se ha procesado el cronograma actual, ¿Desea procesarlo ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        If Rpta = vbYes Then
            procesarCronograma
            MostrarSegundoTab
            Exit Sub
        Else
            Fg1.Rows = 1
            Set RstDet = Nothing
            Exit Sub
        End If
    End If
    
    Dim A As Integer
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        Agregando = True
        'se dan valores iniciales
        fchiniAux = NulosC(RstDet("fchpro"))
        horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
        horprodAux = NulosC(RstDet("horpro"))
        canproAux = NulosN(RstDet("cantidad"))
        ' valores para la materia prima
        matpriAux = NulosC(RstDet("matpri"))
        iditemAux = NulosC(RstDet("iditem"))
        ' valores para el producto
        desproAux = NulosC(RstDet("despro"))
        idproAux = NulosN(RstDet("idpro"))
        
        For A = 1 To RstDet.RecordCount
            If Fg1.Rows = 2 Then
                If es_matpri Then
                    Fg1.Select 1, 1, 1, 1
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &HC0&
                    Fg1.TextMatrix(1, 1) = fchiniAux
                    
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(2, 2) = matpriAux
                    Fg1.Rows = Fg1.Rows + 1
                    
                    Fg1.Select 3, 3, 3, 3
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &H80000002
                    Fg1.TextMatrix(3, 3) = desproAux
                    Fg1.Rows = Fg1.Rows + 1
                Else
                    Fg1.Select 1, 1, 1, 1
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &HC0&
                    Fg1.TextMatrix(1, 1) = fchiniAux
                    
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.Select 2, 2, 2, 2
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &H80000002
                    Fg1.TextMatrix(2, 2) = desproAux
                    Fg1.Rows = Fg1.Rows + 1
                End If
            End If
            
            'Se calcula el tiempo que demora la tarea
            If NulosN(RstDet("aplpor")) <> 0 Then
                xTiempo = (RstDet("factor") * RstDet("cantidad") * (RstDet("aplpor") / 100)) / RstDet("numper")
            Else
                xTiempo = (RstDet("factor") * RstDet("cantidad")) / RstDet("numper")
            End If
            If xTiempo > 8 Then xTiempo = 8
            'If xTiempo = 0 Then xTiempo = 2
            
            xTiempo = Format(Int(xTiempo), "00") & ":" & Format(((xTiempo * 60) Mod 60), "00")
            
            f_destar = Fg1.Rows - 1
            
            If es_matpri Then
                'si es la misma fecha
                If RstDet("fchpro") = fchiniAux Then
                    'si es la misma materia prima
                    If RstDet("iditem") = iditemAux Then
                        'si es el mismo producto
                        If RstDet("idpro") = idproAux Then
                            If RstDet("horpro") = horprodAux Then
                                'se llena la tarea
                                llenarTarea RstDet, Fg1, f_destar, 4, horiniAux, xTiempo, canproAux
                            Else
                                horprodAux = NulosC(RstDet("horpro"))
                                horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                                canproAux = NulosN(RstDet("cantidad"))
                                'se llena el producto
                                desproAux = NulosC(RstDet("despro"))
                                idproAux = NulosN(RstDet("idpro"))
                                f_despro = f_destar
                                Fg1.Select f_despro, 3, f_despro, 3
                                Fg1.FillStyle = flexFillRepeat
                                Fg1.CellForeColor = &H80000002
                                Fg1.TextMatrix(f_despro, 3) = desproAux
                                Fg1.Rows = Fg1.Rows + 1
                                'se llena la tarea
                                f_destar = Fg1.Rows - 1
                                llenarTarea RstDet, Fg1, f_destar, 4, horiniAux, xTiempo, canproAux
                            End If
                        Else
                            horprodAux = NulosC(RstDet("horpro"))
                            horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                            canproAux = NulosN(RstDet("cantidad"))
                            'se llena el producto
                            desproAux = NulosC(RstDet("despro"))
                            idproAux = NulosN(RstDet("idpro"))
                            f_despro = f_destar
                            Fg1.Select f_despro, 3, f_despro, 3
                            Fg1.FillStyle = flexFillRepeat
                            Fg1.CellForeColor = &H80000002
                            Fg1.TextMatrix(f_despro, 3) = desproAux
                            Fg1.Rows = Fg1.Rows + 1
                            'se llena la tarea
                            f_destar = Fg1.Rows - 1
                            llenarTarea RstDet, Fg1, f_destar, 4, horiniAux, xTiempo, canproAux
                        End If
                    Else
                        horprodAux = NulosC(RstDet("horpro"))
                        horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                        canproAux = NulosN(RstDet("cantidad"))
                        'se llena la materia prima
                        matpriAux = NulosC(RstDet("matpri"))
                        iditemAux = NulosN(RstDet("iditem"))
                        f_matpri = f_destar
                        Fg1.TextMatrix(f_matpri, 2) = matpriAux
                        Fg1.Rows = Fg1.Rows + 1
                        'se llena el producto
                        f_destar = Fg1.Rows - 1
                        desproAux = NulosC(RstDet("despro"))
                        idproAux = NulosN(RstDet("idpro"))
                        f_despro = f_destar
                        Fg1.Select f_despro, 3, f_despro, 3
                        Fg1.FillStyle = flexFillRepeat
                        Fg1.CellForeColor = &H80000002
                        Fg1.TextMatrix(f_despro, 3) = desproAux
                        Fg1.Rows = Fg1.Rows + 1
                        'se llena la tarea
                        f_destar = Fg1.Rows - 1
                        llenarTarea RstDet, Fg1, f_destar, 4, horiniAux, xTiempo, canproAux
                    End If
                Else
                    horprodAux = NulosC(RstDet("horpro"))
                    horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                    canproAux = NulosN(RstDet("cantidad"))
                    'se llena la fecha
                    fchiniAux = NulosC(RstDet("fchpro"))
                    f_fchini = f_destar
                    Fg1.Select f_fchini, 1, f_fchini, 1
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &HC0&
                    Fg1.TextMatrix(f_fchini, 1) = fchiniAux
                    Fg1.Rows = Fg1.Rows + 1
                    'se llena la materia prima
                    f_destar = Fg1.Rows - 1
                    matpriAux = NulosC(RstDet("matpri"))
                    iditemAux = NulosN(RstDet("iditem"))
                    f_matpri = f_destar
                    Fg1.TextMatrix(f_matpri, 2) = matpriAux
                    Fg1.Rows = Fg1.Rows + 1
                    'se llena el producto
                    f_destar = Fg1.Rows - 1
                    desproAux = NulosC(RstDet("despro"))
                    idproAux = NulosN(RstDet("idpro"))
                    f_despro = f_destar
                    Fg1.Select f_despro, 3, f_despro, 3
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &H80000002
                    Fg1.TextMatrix(f_despro, 3) = desproAux
                    Fg1.Rows = Fg1.Rows + 1
                    'se llena la tarea
                    f_destar = Fg1.Rows - 1
                    llenarTarea RstDet, Fg1, f_destar, 4, horiniAux, xTiempo, canproAux
                End If
            Else
                'si es la misma fecha
                If RstDet("fchpro") = fchiniAux Then
                    'si es el mismo producto
                    If RstDet("idpro") = idproAux Then
                        If RstDet("horpro") = horprodAux Then
                            'se llena la tarea
                            llenarTarea RstDet, Fg1, f_destar, 3, horiniAux, xTiempo, canproAux
                        Else
                            horprodAux = NulosC(RstDet("horpro"))
                            horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                            canproAux = NulosN(RstDet("cantidad"))
                            'se llena el producto
                            desproAux = NulosC(RstDet("despro"))
                            idproAux = NulosN(RstDet("idpro"))
                            f_despro = f_destar
                            Fg1.Select f_despro, 2, f_despro, 2
                            Fg1.FillStyle = flexFillRepeat
                            Fg1.CellForeColor = &H80000002
                            Fg1.TextMatrix(f_despro, 2) = desproAux
                            Fg1.Rows = Fg1.Rows + 1
                            'se llena la tarea
                            f_destar = Fg1.Rows - 1
                            llenarTarea RstDet, Fg1, f_destar, 3, horiniAux, xTiempo, canproAux
                        End If
                    Else
                        horprodAux = NulosC(RstDet("horpro"))
                        horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                        canproAux = NulosN(RstDet("cantidad"))
                        'se llena el producto
                        desproAux = NulosC(RstDet("despro"))
                        idproAux = NulosN(RstDet("idpro"))
                        f_despro = f_destar
                        Fg1.Select f_despro, 2, f_despro, 2
                        Fg1.FillStyle = flexFillRepeat
                        Fg1.CellForeColor = &H80000002
                        Fg1.TextMatrix(f_despro, 2) = desproAux
                        Fg1.Rows = Fg1.Rows + 1
                        'se llena la tarea
                        f_destar = Fg1.Rows - 1
                        llenarTarea RstDet, Fg1, f_destar, 3, horiniAux, xTiempo, canproAux
                    End If
                Else
                    horprodAux = NulosC(RstDet("horpro"))
                    horiniAux = NulosC(Format(RstDet("horinitar"), "hh:mm"))
                    canproAux = NulosN(RstDet("cantidad"))
                    'se llena la fecha
                    fchiniAux = NulosC(RstDet("fchpro"))
                    f_fchini = f_destar
                    Fg1.Select f_fchini, 1, f_fchini, 1
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &HC0&
                    Fg1.TextMatrix(f_fchini, 1) = fchiniAux
                    Fg1.Rows = Fg1.Rows + 1
                    'se llena el producto
                    f_destar = Fg1.Rows - 1
                    desproAux = NulosC(RstDet("despro"))
                    idproAux = NulosN(RstDet("idpro"))
                    f_despro = f_destar
                    Fg1.Select f_despro, 2, f_despro, 2
                    Fg1.FillStyle = flexFillRepeat
                    Fg1.CellForeColor = &H80000002
                    Fg1.TextMatrix(f_despro, 2) = desproAux
                    Fg1.Rows = Fg1.Rows + 1
                    'se llena la tarea
                    f_destar = Fg1.Rows - 1
                    llenarTarea RstDet, Fg1, f_destar, 3, horiniAux, xTiempo, canproAux
                End If
            End If
            RstDet.MoveNext
        Next A
        Agregando = False
        Fg1.Select 1, 7
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then
            MostrarSegundoTab
            Exit Sub
        End If
    End If
End Sub

Function escribirCronograma() As Boolean
    Dim xId As Double
    Dim xCampos(4, 4) As String
    Dim xCampos2(14, 4) As String
    Dim A As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    xId = RstLis("id")
    
    xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & xId & ""
    
    For A = 1 To Fg2.Rows - 1
        xCampos2(0, 0) = "id":           xCampos2(0, 1) = xId:            xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":     xCampos2(0, 4) = "":
        xCampos2(1, 0) = "fchpro":       xCampos2(1, 1) = Fg2.TextMatrix(A, 9):    xCampos2(1, 2) = "S":    xCampos2(1, 3) = "F":     xCampos2(1, 4) = "No ha especificado la fecha de programacion"
        xCampos2(2, 0) = "horpro":       xCampos2(2, 1) = Fg2.TextMatrix(A, 10):   xCampos2(2, 2) = "S":    xCampos2(2, 3) = "F":     xCampos2(2, 4) = "No ha especificado la hora de recepcion"
        xCampos2(3, 0) = "iditem":       xCampos2(3, 1) = Fg2.TextMatrix(A, 1):    xCampos2(3, 2) = "S":    xCampos2(3, 3) = "N":     xCampos2(3, 4) = "No ha especificado la materia prima"
        xCampos2(4, 0) = "idpro":        xCampos2(4, 1) = Fg2.TextMatrix(A, 2):    xCampos2(4, 2) = "S":    xCampos2(4, 3) = "N":     xCampos2(4, 4) = "No ha especificado el producto"
        xCampos2(5, 0) = "idtar":        xCampos2(5, 1) = Fg2.TextMatrix(A, 3):    xCampos2(5, 2) = "S":    xCampos2(5, 3) = "N":     xCampos2(5, 4) = "No ha especificado la tarea"
        xCampos2(6, 0) = "cantidad":     xCampos2(6, 1) = Fg2.TextMatrix(A, 13):   xCampos2(6, 2) = "S":    xCampos2(6, 3) = "N":     xCampos2(6, 4) = "No ha especificado la cantidad"
        xCampos2(7, 0) = "factor":       xCampos2(7, 1) = Fg2.TextMatrix(A, 4):    xCampos2(7, 2) = "S":    xCampos2(7, 3) = "N":     xCampos2(7, 4) = "No ha especificado el factor"
        xCampos2(8, 0) = "costokg":      xCampos2(8, 1) = 0:                       xCampos2(8, 2) = "S":    xCampos2(8, 3) = "N":     xCampos2(8, 4) = "No ha especificado el costo por kilo"
        xCampos2(9, 0) = "numper":       xCampos2(9, 1) = Fg2.TextMatrix(A, 11):   xCampos2(9, 2) = "S":    xCampos2(9, 3) = "N":     xCampos2(9, 4) = "No ha especificado el numero de personas"
        xCampos2(10, 0) = "horarr":      xCampos2(10, 1) = Fg2.TextMatrix(A, 12):  xCampos2(10, 2) = "S":    xCampos2(10, 3) = "F":     xCampos2(10, 4) = "No ha especificado tiempo en que enpieza cada tarea"
        xCampos2(11, 0) = "aplpor":      xCampos2(11, 1) = Fg2.TextMatrix(A, 14):  xCampos2(11, 2) = "S":    xCampos2(11, 3) = "N":     xCampos2(11, 4) = "No ha especificado el porcentaje de rendimiento"
        xCampos2(12, 0) = "orden":       xCampos2(12, 1) = Fg2.TextMatrix(A, 5):   xCampos2(12, 2) = "S":    xCampos2(12, 3) = "N":     xCampos2(12, 4) = "No ha especificado el orden"
        xCampos2(13, 0) = "horinitar":   xCampos2(13, 1) = Fg2.TextMatrix(A, 17):  xCampos2(13, 2) = "S":    xCampos2(13, 3) = "F":     xCampos2(13, 4) = "No ha especificado la hora de inicio de la tarea"
        xCampos2(14, 0) = "fchini":      xCampos2(14, 1) = Fg2.TextMatrix(A, 19):  xCampos2(14, 2) = "S":    xCampos2(14, 3) = "F":     xCampos2(14, 4) = "No ha especificado la fecha de inicio"
        
        If EscribirNuevoRegistro(xCampos2, "pro_cronogramatarea", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    MsgBox "El cronograma se proceso con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    escribirCronograma = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente error : " & Trim(Err.Description)
    escribirCronograma = False
End Function

Function Grabar() As Boolean
    Dim xId As Double
    Dim xCampos(4, 4) As String
    Dim xCampos2(14, 4) As String
    Dim A As Integer
    
    Dim Rst As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim cSQL As String
    Dim idproAux As Integer
    Dim horproAux As String
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    xId = RstLis("id")
    xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & xId & ""
    RST_Busq RstDet, "SELECT * FROM pro_cronogramatarea", xCon
    
    idproAux = NulosN(Fg1.TextMatrix(1, Fg1.Cols - 4))
    horproAux = NulosC(Fg1.TextMatrix(1, Fg1.Cols - 12))
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, Fg1.Cols - 3) <> "" Then
            'Se verfica que sea el mismo producto
            If idproAux <> Fg1.TextMatrix(A, Fg1.Cols - 4) Then
                idproAux = Fg1.TextMatrix(A, Fg1.Cols - 4)
                horproAux = Fg1.TextMatrix(A, Fg1.Cols - 12)
            End If
            RstDet.AddNew
    
            RstDet("id") = xId
            RstDet("fchpro") = Fg1.TextMatrix(A, Fg1.Cols - 6)
            RstDet("horpro") = horproAux
            RstDet("iditem") = Fg1.TextMatrix(A, Fg1.Cols - 5)
            RstDet("idpro") = Fg1.TextMatrix(A, Fg1.Cols - 4)
            RstDet("idtar") = Fg1.TextMatrix(A, Fg1.Cols - 3)
            RstDet("cantidad") = Fg1.TextMatrix(A, Fg1.Cols - 9)
            RstDet("factor") = Fg1.TextMatrix(A, Fg1.Cols - 2)
            RstDet("costokg") = 0
            RstDet("numper") = NulosN(Fg1.TextMatrix(A, Fg1.Cols - 10))
            RstDet("horarr") = Fg1.TextMatrix(A, 5)
            RstDet("aplpor") = Fg1.TextMatrix(A, Fg1.Cols - 8)
            RstDet("orden") = Fg1.TextMatrix(A, Fg1.Cols - 1)
            RstDet("horinitar") = Fg1.TextMatrix(A, Fg1.Cols - 12)
            RstDet("fchini") = Fg1.TextMatrix(A, Fg1.Cols - 6)
            
            RstDet.Update
        Else
            idproAux = 0
            horproAux = ""
        End If
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    MsgBox "El cronograma se proceso con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente error : " & Trim(Err.Description)
    Grabar = False
End Function

Sub procesarCronograma()
    Dim xRs As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim SQLCad As String
    
    cSQL = "SELECT * FROM pro_cronogramatarea WHERE id = " & RstLis("id") & ""

    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        If NulosN(RstLis("idtippro")) = 3 Then
            ' SI SE ESTAN PROCESANDO PRODUCTOS
            cSQL = "SELECT pro_cronogramadet.id, pro_cronogramadet.fchpro, pro_cronogramadet.Horpro, 0 AS iditem, pro_cronogramadet.iditem AS idpro, " _
                    & " pro_receta.descripcion AS nomrec, '' AS nommatpri, alm_inventario.descripcion AS nompro, pro_recetatar.idtar, pro_tareas.codigo AS codtar, " _
                    & " pro_tareas.descripcion AS destar, pro_cronogramadet.cantidad, pro_recetatar.factor, pro_recetatar.costokg, pro_recetatar.numper, pro_recetatar.horarr, " _
                    & " pro_recetatar.aplpor, pro_recetatar.orden, IIf(pro_recetatar.aplpor=0,pro_recetatar.factor*pro_cronogramadet.cantidad,(pro_recetatar.factor*pro_cronogramadet.cantidad)*(pro_recetatar.aplpor/100)) AS tiempoesttotal, " _
                    & " [tiempoesttotal]/[numper] AS tiempoesttotalper, [Horpro]+pro_recetatar.horarr AS horinitar, Format(Int(([tiempoesttotal]*60)/60),'00') & ':' & Format(([tiempoesttotal]*60) Mod 60,'00') AS tiempoesttotalhor, " _
                    & " Format(Int(([tiempoesttotalper]*60)/60),'00') & ':' & Format(([tiempoesttotalper]*60) Mod 60,'00') AS tiempoesttotalhorperr, pro_cronogramadet.fchpro AS fchini " _
                + vbCr + "FROM pro_tareas RIGHT JOIN (((pro_cronogramadet LEFT JOIN pro_receta ON pro_cronogramadet.iditem = pro_receta.iditem) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar" _
                + vbCr + "Where (((pro_cronogramadet.id) = " & RstLis("id") & ") AND ((pro_cronogramadet.cantidad)<>0) And ((pro_recetatar.factor) <> 0) And ((pro_recetatar.numper) <> 0) And ((pro_receta.prirec) >= 0)) " _
                + vbCr + "ORDER BY pro_cronogramadet.fchpro, pro_cronogramadet.Horpro, pro_receta.descripcion, pro_recetatar.orden"
        Else
            ' SI SE ESTA PROCESANDO MATERIA PRIMA
            cSQL = "SELECT pro_cronogramadetprod.id, pro_cronogramadetprod.fchpro, pro_cronogramadetprod.horpro, pro_cronogramadetprod.iditem, pro_cronogramadetprod.idpro, " _
                    & " pro_receta.descripcion AS nomrec, alm_inventario_1.descripcion AS nommatpri, alm_inventario.descripcion AS nompro, pro_recetatar.idtar, pro_tareas.codigo AS codtar, " _
                    & " pro_tareas.descripcion AS destar, pro_cronogramadetprod.cantidad, pro_recetatar.factor, pro_recetatar.costokg, pro_recetatar.numper, pro_recetatar.horarr, " _
                    & " pro_recetatar.aplpor, pro_recetatar.orden, IIf([pro_recetatar].[aplpor]=0,[pro_recetatar].[factor]*[pro_cronogramadetprod].[cantidad],([pro_recetatar].[factor]*[pro_cronogramadetprod].[cantidad])*([pro_recetatar].[aplpor]/100)) AS tiempoesttotal, " _
                    & " [tiempoesttotal]/[numper] AS tiempoesttotalper, [pro_cronogramadetprod].[Horpro]+[pro_recetatar].[horarr] AS horinitar, " _
                    & " Format(Int(([tiempoesttotal]*60)/60),'00') & ':' & Format(([tiempoesttotal]*60) Mod 60,'00') AS tiempoesttotalhor, " _
                    & " Format(Int(([tiempoesttotalper]*60)/60),'00') & ':' & Format(([tiempoesttotalper]*60) Mod 60,'00') AS tiempoesttotalhorperr, pro_cronogramadetprod.fchpro AS fchini, " _
                    & " pro_receta.prirec " _
                + vbCr + "FROM ((((alm_inventario AS alm_inventario_1 RIGHT JOIN pro_cronogramadetprod ON alm_inventario_1.id = pro_cronogramadetprod.iditem) LEFT JOIN pro_receta ON pro_cronogramadetprod.idpro = pro_receta.iditem) LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) LEFT JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id " _
                + vbCr + "Where (((pro_cronogramadetprod.id) = " & RstLis("id") & ") AND ((pro_cronogramadetprod.cantidad)<>0) And ((pro_recetatar.factor) <> 0) And ((pro_recetatar.numper) <> 0) And ((pro_receta.prirec) >= 0)) " _
                + vbCr + "ORDER BY pro_cronogramadetprod.fchpro, pro_cronogramadetprod.horpro, alm_inventario.descripcion, pro_recetatar.orden"
        End If
        
        RST_Busq RstTar, cSQL, xCon
        
        Fg2.Rows = 1
        Dim A As Integer
        If RstTar.RecordCount <> 0 Then
            RstTar.MoveFirst
            Agregando = True
            Dim xCadLlave As String
            Dim xCadLlave2 As String
            Dim xColor As Long
            Dim xPintar As Boolean
            
            xCadLlave = Format(RstTar("iditem"), "0000") & Format(RstTar("idpro"), "0000") & Format(RstTar("fchpro"), "dd/mm/yy") & Format(RstTar("horpro"), "hh:mm")
            xColor = &HE0FEFE
            xPintar = True
            
            For A = 1 To RstTar.RecordCount
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(A, 1) = RstTar("iditem")
                Fg2.TextMatrix(A, 2) = RstTar("idpro")
                Fg2.TextMatrix(A, 3) = RstTar("idtar")
                Fg2.TextMatrix(A, 4) = RstTar("factor")
                Fg2.TextMatrix(A, 5) = RstTar("orden")
                Fg2.TextMatrix(A, 6) = RstTar("nommatpri")
                Fg2.TextMatrix(A, 7) = RstTar("nompro")
                Fg2.TextMatrix(A, 8) = RstTar("destar")
                Fg2.TextMatrix(A, 9) = Format(RstTar("fchpro"), "dd/mm/yy")
                Fg2.TextMatrix(A, 10) = Format(RstTar("horpro"), "hh:mm")
                Fg2.TextMatrix(A, 11) = Format(RstTar("numper"), "00")
                Fg2.TextMatrix(A, 12) = Format(RstTar("horarr"), "hh:mm")
                Fg2.TextMatrix(A, 13) = Format(RstTar("cantidad"), "0.00")
                Fg2.TextMatrix(A, 14) = Format(RstTar("aplpor"), "0.00")
                
                Fg2.TextMatrix(A, 19) = Format(RstTar("fchini"), "dd/mm/yy")
                
                Dim xTiempo As Double
                Dim xHorEst As String
                xTiempo = 1
                
                If NulosN(RstTar("aplpor")) <> 0 Then
                    Fg2.TextMatrix(A, 15) = (RstTar("cantidad") * ((RstTar("aplpor") / 100)))
                    Fg2.TextMatrix(A, 15) = Format(Fg2.TextMatrix(A, 15), "0.00")
                    
                    xTiempo = (RstTar("factor") * Fg2.TextMatrix(A, 15)) / RstTar("numper")
                Else
                    Fg2.TextMatrix(A, 15) = Format(RstTar("cantidad"), "0.00")
                    xTiempo = (RstTar("factor") * RstTar("cantidad")) / RstTar("numper")
                End If
                xHorEst = ""
                xHorEst = Format(Int(xTiempo), "00")
                xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
                Fg2.TextMatrix(A, 16) = xHorEst
                
                Fg2.TextMatrix(A, 17) = Format(RstTar("horinitar"), "hh:mm")
                
                If Val(Mid(xHorEst, 1, 2)) <= 8 Then
                    Fg2.TextMatrix(A, 18) = Format(CDate(xHorEst) + CDate(Fg2.TextMatrix(A, 17)), "HH:MM")
                Else
                    Fg2.TextMatrix(A, 18) = ""
                End If
                If xPintar = True Then
                    GRID_COLOR_FONDO Fg2, Fg2.Rows - 1, 1, Fg2.Rows - 1, 19, xColor, flexFillRepeat
                End If
                
                RstTar.MoveNext
                If RstTar.EOF = True Then
                    Exit For
                End If
                xCadLlave2 = Format(RstTar("iditem"), "0000") & Format(RstTar("idpro"), "0000") & Format(RstTar("fchpro"), "dd/mm/yy") & Format(RstTar("horpro"), "hh:mm")
                If xCadLlave2 <> xCadLlave Then
                    xCadLlave = xCadLlave2
                    xPintar = Not xPintar
                End If
            Next A
            
            Agregando = False
        End If
    Else
        MostrarSegundoTab
    End If
    
    xHorIni = Time
    
    If escribirCronograma Then RstLis.Requery: Dg1.Refresh
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    
    If RstLis.RecordCount = 0 Then
        MsgBox "No hay registro para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar las tareas del cronograma seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & RstLis("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLis("id") & " AND idform = " & IdMenuActivo
                
        MsgBox "El cronograma de tareas se elimino con exito"
        RstLis.Requery
        Dg1.Refresh
    End If
End Sub

Sub Modificar()
    Dim columna_duracion As Integer
    Dim columna_horIni As Integer
    Dim columna_horFin As Integer
    Dim columna_numPer As Integer
    Dim columna_cantidad As Integer
    
    columna_duracion = Fg1.Cols - 13
    columna_horIni = Fg1.Cols - 12
    columna_horFin = Fg1.Cols - 11
    columna_numPer = Fg1.Cols - 10
    columna_cantidad = Fg1.Cols - 9
    
    QueHace = 2
    Bloquea
    Label1.Caption = "Modificando Programacion de Tareas"
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Fg1.SelectionMode = flexSelectionFree
    Fg1.AutoSearch = flexSearchNone
    Fg1.Editable = flexEDKbdMouse
    
    Fg1.Select 1, columna_horIni, Fg1.Rows - 1, columna_horIni
    Fg1.CellForeColor = &HFF&       '&H8000000F
    Fg1.Select 1, columna_cantidad, Fg1.Rows - 1, Fg1.Cols - 1
    Fg1.CellForeColor = &HFF&
    
    Fg1.Select 1, 1, 1, Fg1.Cols - 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLis.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        RstLis.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 14 Then ExportarExcel
    
    If Button.Index = 16 Then
        RstLis.Close
        Set RstLis = Nothing
        Unload Me
    End If
End Sub

Sub ExportarExcel()
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
        .Cells(1, 2) = "Cronograma de Tareas"
        .Range("B1", "L1").Merge
        .Cells(1, 2).HorizontalAlignment = xlHAlignCenterAcrossSelection
        .Cells(1, 2).Font.Bold = True
        .Cells(1, 2).Rows(1).Font.Size = 12
        
        .Cells(2, 2) = "Fecha de Inicio: "
        .Cells(2, 2).Font.Bold = True
        .Cells(2, 3) = CDate(RstLis("fchini"))
        .Cells(3, 2) = "Fecha de Fin: "
        .Cells(3, 2).Font.Bold = True
        .Cells(3, 3) = CDate(RstLis("fchfin"))
        xFilas = 5
        For A = 0 To Fg1.Rows - 1
            For B = 1 To Fg1.Cols - 1
                If B < Fg1.Cols - 6 Then
                    If A = 0 Then
                        .Cells(xFilas, B + 1).Font.Bold = True
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    Else
                        If B <= 7 Then
                            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        Else
                            If Fg1.TextMatrix(A, B) <> "" Then .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                        End If
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
