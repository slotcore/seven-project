VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmEmisionPlanilla 
   Caption         =   "Planillas - Emision de Planillas"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4590
      Top             =   0
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
            Picture         =   "FrmEmisionPlanilla.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmisionPlanilla.frx":1EA4
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
      Width           =   11700
      _ExtentX        =   20638
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
            Object.ToolTipText     =   "Reportes"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s1"
                  Object.Tag             =   "1"
                  Text            =   "Lista de Productos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s2"
                  Object.Tag             =   "2"
                  Text            =   "Lista de Precios"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s3"
                  Object.Tag             =   "3"
                  Text            =   "Productos sin Stock"
               EndProperty
            EndProperty
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
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   11700
      _cx             =   20637
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
      BackTabColor    =   8421504
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   29
         Top             =   375
         Width           =   11610
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   30
            Top             =   390
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Planilla"
            Columns(0).DataField=   "numpla"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tipo Planilla"
            Columns(1).DataField=   "destippla"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Proceso"
            Columns(2).DataField=   "fchpro"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Inicio"
            Columns(3).DataField=   "fchini"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Final"
            Columns(4).DataField=   "fchfin"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tot. Basico"
            Columns(5).DataField=   "totbas"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Tot. Dsct."
            Columns(6).DataField=   "totdes"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Apo. Emp."
            Columns(7).DataField=   "totapoemp"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2328"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2249"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2143"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2064"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2275"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2196"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2196"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2117"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2064"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1984"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE7FEFC&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H800000&,.bold=-1"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.alignment=2"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14,.alignment=2"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Planillas"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   12345
         TabIndex        =   2
         Top             =   375
         Width           =   11610
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   390
            Left            =   8460
            TabIndex        =   33
            Top             =   1065
            Width           =   1005
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   240
            Left            =   9600
            TabIndex        =   32
            Top             =   900
            Width           =   1320
         End
         Begin VB.TextBox TxtNumPla 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "TxtNumPla"
            Top             =   405
            Width           =   1230
         End
         Begin VB.TextBox TxtBusTipPla 
            Height          =   300
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "TxtBusTipPla"
            Top             =   705
            Width           =   1230
         End
         Begin VB.CommandButton CmdBusGrupo 
            Height          =   225
            Left            =   5100
            Picture         =   "FrmEmisionPlanilla.frx":23EC
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   735
            Width           =   240
         End
         Begin VB.TextBox TxtTotBas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   7110
            TabIndex        =   6
            Text            =   "TxtTotBas"
            Top             =   6315
            Width           =   1095
         End
         Begin VB.TextBox TxtTotDsct 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   8235
            TabIndex        =   5
            Text            =   "TxtTotDsct"
            Top             =   6315
            Width           =   1095
         End
         Begin VB.TextBox TxtTotApo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10485
            TabIndex        =   4
            Text            =   "TxtTotApo"
            Top             =   6315
            Width           =   1095
         End
         Begin VB.TextBox TxtTotPla 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   9360
            TabIndex        =   3
            Text            =   "TxtTotPla"
            Top             =   6315
            Width           =   1095
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4350
            Left            =   45
            TabIndex        =   7
            Top             =   1665
            Width           =   11535
            _cx             =   20346
            _cy             =   7673
            _ConvInfo       =   1
            Appearance      =   0
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
            Cols            =   3
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEmisionPlanilla.frx":251E
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
         Begin AspaTextBoxFecha.TextBoxFecha txtFchPro 
            Height          =   300
            Left            =   1320
            TabIndex        =   10
            Top             =   705
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "10/10/2005"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   1320
            TabIndex        =   12
            Top             =   1020
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "10/10/2005"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   4140
            TabIndex        =   13
            Top             =   1020
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "10/10/2005"
         End
         Begin VB.Frame Frame3 
            Height          =   690
            Left            =   45
            TabIndex        =   14
            Top             =   5970
            Width           =   6960
            Begin VB.Shape Shape1 
               BackColor       =   &H00E7FEFC&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   135
               Top             =   255
               Width           =   555
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   1995
               Top             =   255
               Width           =   555
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   4005
               Top             =   255
               Width           =   555
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Ingresos"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   780
               TabIndex        =   17
               Top             =   270
               Width           =   840
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Descuentos"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2670
               TabIndex        =   16
               Top             =   270
               Width           =   1050
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Aportaciones del Emp."
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4665
               TabIndex        =   15
               Top             =   270
               Width           =   2205
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Planilla"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   90
            TabIndex        =   28
            Top             =   60
            Width           =   11400
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Planilla"
            Height          =   195
            Left            =   60
            TabIndex        =   27
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Proceso"
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   750
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   25
            Top             =   1050
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Final"
            Height          =   195
            Left            =   3255
            TabIndex        =   24
            Top             =   1050
            Width           =   690
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nomina de Empleados"
            Height          =   225
            Left            =   60
            TabIndex        =   23
            Top             =   1410
            Width           =   1590
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Planilla"
            Height          =   195
            Left            =   3090
            TabIndex        =   22
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Basico"
            Height          =   195
            Left            =   7110
            TabIndex        =   21
            Top             =   6090
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Dsct."
            Height          =   195
            Left            =   8280
            TabIndex        =   20
            Top             =   6090
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tot. Apor. Emp"
            Height          =   195
            Left            =   10500
            TabIndex        =   19
            Top             =   6090
            Width           =   1065
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Total Planilla"
            Height          =   195
            Left            =   9375
            TabIndex        =   18
            Top             =   6090
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "FrmEmisionPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CargarPlanillas()
    Dim RstPla As New ADODB.Recordset
    Dim A, B As Integer
    Dim xCol, xFil, UltCol As Integer
    Dim xColPos1, xColPos2 As Integer
    
    RST_Busq RstPla, "TRANSFORM Sum(IIf(mae_concepingresosdet!tipo=-1,pla_empleadosapo!importe,mae_concepingresosdet!valref*(mae_concepingresosdet!importe/100))) AS importe" _
        & " SELECT pla_empleados.id, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS nomemp FROM pla_empleados " _
        & " INNER JOIN (mae_concepingresosdet RIGHT JOIN pla_empleadosapo ON mae_concepingresosdet.id = pla_empleadosapo.idconcep) ON pla_empleados.id = pla_empleadosapo.idemp " _
        & " GROUP BY pla_empleados.id, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] PIVOT pla_empleadosapo.idconcep", xCon

    Fg1.Rows = 2
    
    'escribimos las cabecera de los ingresos
    xColPos1 = Fg1.Cols - 1
    For B = 2 To RstPla.Fields.Count - 1
        Fg1.Cols = Fg1.Cols + 1
        
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "INGRESOS"
        Fg1.TextMatrix(1, Fg1.Cols - 1) = "I-" + Trim(RstPla.Fields(B).Name)
    Next B
    xColPos2 = Fg1.Cols - 1
    
    
    xFil = 2
    
    RstPla.MoveFirst
    Fg1.Rows = Fg1.Rows + 1
    For A = 1 To RstPla.RecordCount
        Fg1.TextMatrix(xFil, 1) = RstPla("nomemp")
        Fg1.TextMatrix(xFil, 2) = RstPla("id")
        
        xCol = 3
        For B = 2 To RstPla.Fields.Count - 1
            Fg1.TextMatrix(xFil, xCol) = Format(NulosN(RstPla.Fields(B).Value), "0.00")
            xCol = xCol + 1
        Next B
        
        RstPla.MoveNext
        If RstPla.EOF = True Then
            Exit For
        End If
        
        Fg1.Rows = Fg1.Rows + 1
        xFil = xFil + 1
    Next A
    'pintamos  los ingreso del color que corresponde
    With Fg1
        .Select 2, xColPos1, Fg1.Rows - 1, xColPos2
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE7FEFC
    End With
    UNIR_CELDAS Fg1, 0, 1, 1, 1, "Empleado/Trabajador", flexAlignCenterCenter, False
    Fg1.MergeCells = flexMergeFixedOnly
    
    UNIR_CELDAS Fg1, 0, CLng(xColPos1), 0, CLng(xColPos2), "INGRES0S", flexAlignCenterCenter, True
    Fg1.MergeCells = flexMergeFixedOnly
    
    
    UltCol = Fg1.Cols - 1
    '-------------------------------------------------------------------------
    'escribimos las cabecera de los descuentos que se aplicaran al trabajador
    RST_Busq RstPla, "TRANSFORM Sum(IIf(mae_concepingresosdet!tipo=-1,pla_empleadosapo!importe*(mae_aportes!porcentaje/100)," _
        & " ((mae_concepingresosdet!valref*(mae_concepingresosdet!importe/100))*(mae_aportes!porcentaje/100)))) AS importe2 " _
        & " SELECT pla_empleados.id, pla_empleados.apepat, pla_empleados.nom FROM pla_empleados RIGHT JOIN (mae_aportes RIGHT " _
        & " JOIN (mae_concepingresosdet RIGHT JOIN (mae_concepingresosdetapo LEFT JOIN pla_empleadosapo ON " _
        & " mae_concepingresosdetapo.idconcep = pla_empleadosapo.idconcep) ON mae_concepingresosdet.id = pla_empleadosapo.idconcep) " _
        & " ON mae_aportes.id = mae_concepingresosdetapo.idaporte) ON pla_empleados.id = pla_empleadosapo.idemp Where (((mae_aportes.tipo) = -1)) " _
        & " GROUP BY pla_empleados.id, pla_empleados.apepat, pla_empleados.nom, mae_aportes.tipo PIVOT mae_aportes.id", xCon

    'AGREGAMOS LAS COLUMNAS PARA LOS DESCUENTOS
    xColPos1 = Fg1.Cols
    For B = 3 To RstPla.Fields.Count - 1
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "DESCUENTOS"
        Fg1.TextMatrix(1, Fg1.Cols - 1) = "D-" + Trim(RstPla.Fields(B).Name)
    Next B
    xColPos2 = Fg1.Cols - 1
    
    'ESCRIBIMOS LOS DESCUENTOS QUE HACE TRABAJADOR
    RstPla.MoveFirst
    xFil = 2
    
    For A = 1 To RstPla.RecordCount
        xCol = UltCol + 1
        For B = 3 To RstPla.Fields.Count - 1
            Fg1.TextMatrix(xFil, xCol) = Format(NulosN(RstPla.Fields(B).Value), "0.00")
            xCol = xCol + 1
        Next B
        
        RstPla.MoveNext
        If RstPla.EOF = True Then
            Exit For
        End If
        xFil = xFil + 1
    Next A
    With Fg1
        .Select 2, xColPos1, Fg1.Rows - 1, xColPos2
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0E0FF
    End With
    
    UNIR_CELDAS Fg1, 0, CLng(xColPos1), 0, CLng(xColPos2), "DESCUENTOS", flexAlignCenterCenter, True
    Fg1.MergeCells = flexMergeFixedOnly
    
    
    
    UltCol = Fg1.Cols - 1
    '-------------------------------------------------------------------------
    'escribimos las cabeceras para los aportes del empleador
    RST_Busq RstPla, "TRANSFORM Sum(IIf(mae_concepingresosdet!tipo=-1,pla_empleadosapo!importe*(mae_aportes!porcentaje/100)," _
        & " ((mae_concepingresosdet!valref*(mae_concepingresosdet!importe/100))*(mae_aportes!porcentaje/100)))) AS importe2 " _
        & " SELECT pla_empleados.id, pla_empleados.apepat, pla_empleados.nom FROM pla_empleados RIGHT JOIN (mae_aportes RIGHT " _
        & " JOIN (mae_concepingresosdet RIGHT JOIN (mae_concepingresosdetapo LEFT JOIN pla_empleadosapo ON " _
        & " mae_concepingresosdetapo.idconcep = pla_empleadosapo.idconcep) ON mae_concepingresosdet.id = pla_empleadosapo.idconcep) " _
        & " ON mae_aportes.id = mae_concepingresosdetapo.idaporte) ON pla_empleados.id = pla_empleadosapo.idemp Where (((mae_aportes.tipo) = 0)) " _
        & " GROUP BY pla_empleados.id, pla_empleados.apepat, pla_empleados.nom, mae_aportes.tipo PIVOT mae_aportes.id", xCon

    'AGREGAMOS LAS COLUMNAS PARA LOS APORTES DL EMPLEADOR
    xColPos1 = Fg1.Cols
    For B = 3 To RstPla.Fields.Count - 1
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "APORTES"
        Fg1.TextMatrix(1, Fg1.Cols - 1) = "A" + Trim(RstPla.Fields(B).Name)
    Next B
    xColPos2 = Fg1.Cols - 1
    
    UNIR_CELDAS Fg1, 0, CLng(xColPos1), 0, CLng(xColPos2), "APORTES", flexAlignCenterCenter, True
    Fg1.MergeCells = flexMergeFixedOnly
    
    'ESCRIBIMOS LOS APORTES DEL EMPLEADOR
    RstPla.MoveFirst
    xFil = 2
    
    For A = 1 To RstPla.RecordCount
        xCol = UltCol + 1
        For B = 3 To RstPla.Fields.Count - 1
            Fg1.TextMatrix(xFil, xCol) = Format(NulosN(RstPla.Fields(B).Value), "0.00")
            xCol = xCol + 1
        Next B
        
        RstPla.MoveNext
        If RstPla.EOF = True Then
            Exit For
        End If
        xFil = xFil + 1
    Next A
    
    With Fg1
        .Select 2, xColPos1, Fg1.Rows - 1, xColPos2
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0C0FF
    End With
    
    
    
    'Calculamos el total de la planilla
    Fg1.Cols = Fg1.Cols + 1
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "TOTAL"
    Fg1.TextMatrix(1, Fg1.Cols - 1) = "TOTAL"
    
    Dim xTotal As Double
    
    For A = 2 To Fg1.Rows - 1
        xTotal = 0
        For B = 3 To Fg1.Cols - 2
            If Mid(Fg1.TextMatrix(1, B), 1, 1) = "I" Then
                xTotal = xTotal + NulosN(Fg1.TextMatrix(A, B))
            End If
        
            If Mid(Fg1.TextMatrix(1, B), 1, 1) = "D" Then
                xTotal = xTotal - NulosN(Fg1.TextMatrix(A, B))
            End If
        Next B
        Fg1.TextMatrix(A, Fg1.Cols - 1) = Format(xTotal, "0.00")
    Next A
    
    UNIR_CELDAS Fg1, 0, Fg1.Cols - 1, 1, Fg1.Cols - 1, "TOTAL", flexAlignCenterCenter, False
    Fg1.MergeCells = flexMergeFixedOnly
    
End Sub

Private Sub Command1_Click()
    Fg1.Cols = 3
    Fg1.Cols = 3
    CargarPlanillas
End Sub

Private Sub Command2_Click()
    FrmPrintBoleta.Show
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(2) = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

