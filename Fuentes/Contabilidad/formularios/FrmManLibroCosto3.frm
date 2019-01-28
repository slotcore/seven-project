VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManLibroCosto3 
   Caption         =   "Contabilidad - Libro de Costos de Producción"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   3000
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   21
         Top             =   420
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cancelar = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   4470
         TabIndex        =   24
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Procesando:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   150
         TabIndex        =   23
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "LblProg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3500
         TabIndex        =   22
         Top             =   180
         Width           =   525
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
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
            Picture         =   "FrmManLibroCosto3.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto3.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7590
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11925
      _cx             =   21034
      _cy             =   13388
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
      Appearance      =   2
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7170
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   11835
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6555
            Left            =   30
            TabIndex        =   15
            Top             =   480
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11562
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Mes"
            Columns(1).DataField=   "desmes"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Almacen"
            Columns(2).DataField=   "desalm"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Descripcion"
            Columns(3).DataField=   "descripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Método Valorización"
            Columns(4).DataField=   "desmetval"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2064"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1984"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=5953"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5874"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=5345"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=5265"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=4974"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4895"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
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
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=3"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15,.alignment=3"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(60)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(63)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   "Named:id=36:Selected"
            _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(66)  =   "Named:id=37:Caption"
            _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(68)  =   "Named:id=38:HighlightRow"
            _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(70)  =   "Named:id=39:EvenRow"
            _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(72)  =   "Named:id=40:OddRow"
            _StyleDefs(73)  =   ":id=40,.parent=33"
            _StyleDefs(74)  =   "Named:id=41:RecordSelector"
            _StyleDefs(75)  =   ":id=41,.parent=34"
            _StyleDefs(76)  =   "Named:id=42:FilterBar"
            _StyleDefs(77)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Libro de Costo de Producción"
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
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   14
            Top             =   30
            Width           =   11685
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7170
         Left            =   12570
         TabIndex        =   11
         Top             =   375
         Width           =   11835
         Begin VB.CommandButton cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   1830
            Picture         =   "FrmManLibroCosto3.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   980
            Width           =   240
         End
         Begin VB.TextBox CIFText 
            BackColor       =   &H0080FFFF&
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   3
            Text            =   "CIFText"
            Top             =   1825
            Width           =   2850
         End
         Begin VB.TextBox MODText 
            BackColor       =   &H0080FFFF&
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   2
            Text            =   "MODText"
            Top             =   1540
            Width           =   2850
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Configuración ]"
            ForeColor       =   &H00800000&
            Height          =   1845
            Left            =   5580
            TabIndex        =   25
            Top             =   330
            Width           =   4575
            Begin VB.CheckBox ckProcFal 
               Caption         =   "Procesar Faltantes"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1500
               Width           =   1695
            End
            Begin VB.CheckBox ReprocesarCheck 
               Caption         =   "Reprocesar Todo"
               Height          =   255
               Left            =   120
               TabIndex        =   5
               Top             =   1260
               Width           =   2055
            End
            Begin VB.CommandButton cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   750
               Picture         =   "FrmManLibroCosto3.frx":2C42
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   630
               Width           =   240
            End
            Begin VB.TextBox txtmetval 
               Height          =   300
               Left            =   90
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   4
               Text            =   "txtmetval"
               Top             =   600
               Width           =   915
            End
            Begin VB.Label lblidmetval 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidtipdist"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   3690
               TabIndex        =   39
               Top             =   300
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label Label13 
               Caption         =   "[ Opciones de Proceso ]"
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   960
               Width           =   1815
            End
            Begin VB.Line Line4 
               BorderColor     =   &H8000000C&
               X1              =   120
               X2              =   4470
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label lblmetval 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblmetval"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   1020
               TabIndex        =   28
               Top             =   600
               Width           =   3465
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Conf. de Valorizacion:"
               Height          =   195
               Left            =   90
               TabIndex        =   27
               Top             =   270
               Width           =   1545
            End
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Procesar &MOD/CIF"
            Enabled         =   0   'False
            Height          =   490
            Index           =   3
            Left            =   10230
            TabIndex        =   7
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   1650
            Width           =   1400
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Procesar MP"
            Enabled         =   0   'False
            Height          =   350
            Index           =   2
            Left            =   10230
            TabIndex        =   6
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   1215
            Width           =   1400
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Config. Distrib."
            Height          =   350
            Index           =   1
            Left            =   10230
            TabIndex        =   8
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   855
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Datos de Producción ]"
            ForeColor       =   &H00800000&
            Height          =   4935
            Left            =   0
            TabIndex        =   18
            Top             =   2150
            Width           =   11775
            Begin VB.Frame Frame7 
               Caption         =   "[ Importe de Materia Prima ]"
               ForeColor       =   &H00800000&
               Height          =   3060
               Left            =   100
               TabIndex        =   31
               Top             =   1800
               Width           =   11565
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   2715
                  Index           =   3
                  Left            =   210
                  TabIndex        =   32
                  Top             =   210
                  Width           =   11370
                  _cx             =   20055
                  _cy             =   4789
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
                  Rows            =   2
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManLibroCosto3.frx":2D74
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
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   1455
               Index           =   0
               Left            =   60
               TabIndex        =   19
               Top             =   300
               Width           =   11625
               _cx             =   20505
               _cy             =   2566
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
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   30
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManLibroCosto3.frx":2EB8
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
         Begin VB.TextBox txtdescripcion 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   1
            Text            =   "txtdescripcion"
            Top             =   650
            Width           =   4280
         End
         Begin VB.ComboBox cbMes 
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   350
            Width           =   2865
         End
         Begin VB.TextBox txtalm 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   36
            Text            =   "txtalm"
            Top             =   950
            Width           =   915
         End
         Begin VB.Label lblsumafactor 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "lblsumafactor"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   10260
            TabIndex        =   41
            Top             =   540
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblidalm 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "lblidalm"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   4740
            TabIndex        =   40
            Top             =   420
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   990
            Width           =   615
         End
         Begin VB.Label lblalm 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblalm"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2100
            TabIndex        =   37
            Top             =   950
            Width           =   3345
         End
         Begin VB.Label Label12 
            Caption         =   "[ Valores de Distribucion ]"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1280
            Width           =   1815
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            X1              =   120
            X2              =   5460
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CIF"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1870
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "MOD"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1570
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   380
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   680
            Width           =   840
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Libro de Costo de Producción"
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
            Height          =   255
            Left            =   60
            TabIndex        =   12
            Top             =   30
            Width           =   11685
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Insertar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Ver Receta"
      End
   End
End
Attribute VB_Name = "FrmManLibroCosto3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------VARIABLES DE ESTADO DE FORMULARIO
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim OrigFX As Long
Dim OrigFY As Long
Dim fOrdenLista As Boolean             ' especfica el orden de la lista de la consulta
'***********************************************
'-----------------------VARIABLES DE FORMULARIO
'***********************************************
Dim RstLibro As New ADODB.Recordset
Dim cSQL As String
Dim F As New SistemaLogica.Funciones

Private Sub pProcesarDatos(MESATRABAJAR_ As Integer)
    Dim FWin As New SistemaWindows.SistemaWindowsClass
    Dim FProd As New ProduccionLogica.Funciones
    Dim mMetVal As New ContabilidadEntidad.EConfigVal
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mIndiceTotal As Double
    Dim mCostoMODMovimientoDetalle
    Dim mCostoCIFMovimientoDetalle
    Dim mFechaInicioInventario As Date
    Dim ANIOACTUAL_ As Integer
    Dim MESACTUAL_ As Integer
    Dim PRIMERDIAMES_ As Date
    Dim ULTIMODIAMES_ As Date
    Dim A As Integer
    
On Error GoTo ERROR_
    ' Se validan los valores
    If F.NuloNumeric(MODText.Text) = 0 Then
        F.MostrarMensajeError "El Valor del MOD a distribuir debe ser mayor a cero", "Error"
        Exit Sub
    End If
    If F.NuloNumeric(CIFText.Text) = 0 Then
        F.MostrarMensajeError "El Valor del CIF a distribuir debe ser mayor a cero", "Error"
        Exit Sub
    End If
    ' Se carga la configuracion de la valorizacion
    Set mMetVal.Conexion = xCon
    mMetVal.Fetch NulosN(lblidmetval)
    
    If fg(0).Rows = fg(0).FixedRows Then
        F.MostrarMensajeError "Para procesar MOD/CIF primero debe de cargar los PP", "Error"
        Exit Sub
    End If
    ' Se verifica la columna factor
    mIndiceTotal = F.NuloNumeric(GRID_SUMAR_COL(fg(0), fg(0).ColIndex("FACTOR")))
    If Not F.CompararConCriterio(mIndiceTotal, 100) Then
        F.MostrarMensajeError "Para procesar MOD/CIF el factor de distribucion debe sumar el 100%", "Error"
        Exit Sub
    End If
    If F.NuloString(fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("HORFIN"))) = "TOTAL" Then
        fg(0).Rows = fg(0).Rows - 1
    End If

    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    mFechaInicioInventario = F.FechaInicioMovimientos(F.NuloNumeric(lblidalm.Caption), xCon)
    
    PgBar.Min = 0
    PgBar.Max = fg(0).Rows - 1
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
            
    ' MOD
    If mMetVal.ProcesaMOD Then
        lbl(2).Caption = "APLICANDO MOD"
        With fg(0)
            For A = .FixedRows To .Rows - 1
                DoEvents
                lbl(0).Caption = "Parte: " & F.NuloString(.TextMatrix(A, .ColIndex("NUMPARTE")))
                LblProg.Caption = "Codigo Item: " & F.NuloString(.TextMatrix(A, .ColIndex("RECETA")))
                PgBar.Value = A
                .TopRow = A
                
                mCostoMODMovimientoDetalle = (F.NuloNumeric(MODText.Text) * F.NuloNumeric(.TextMatrix(A, .ColIndex("FACTOR")))) / 100
                .TextMatrix(A, .ColIndex("COSTOMOD")) = Format(mCostoMODMovimientoDetalle, FORMAT_IMPORTEKARDEX)
                                
                ' Actualizamos Valores
                .TextMatrix(A, .ColIndex("COSTOPRIMO")) = Format(F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMP"))) + F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMOD"))), FORMAT_IMPORTEKARDEX)
            Next A
        End With
    End If
    ' CIF
    If mMetVal.ProcesaCIF Then
        lbl(2).Caption = "APLICANDO CIF"
        With fg(0)
            For A = .FixedRows To .Rows - 1
                DoEvents
                lbl(0).Caption = "Parte: " & F.NuloString(.TextMatrix(A, .ColIndex("NUMPARTE")))
                LblProg.Caption = "Codigo Item: " & F.NuloString(.TextMatrix(A, .ColIndex("RECETA")))
                PgBar.Value = A
                .TopRow = A
                
                mCostoCIFMovimientoDetalle = (F.NuloNumeric(CIFText.Text) * F.NuloNumeric(.TextMatrix(A, .ColIndex("FACTOR")))) / 100
                .TextMatrix(A, .ColIndex("COSTOCIF")) = Format(mCostoCIFMovimientoDetalle, FORMAT_IMPORTEKARDEX)
                
                
                ' Grabamos el costo del movimiento
                Dim mIdMovimientoDetalle As Long
                If F.NuloNumeric(.TextMatrix(A, .ColIndex("IDMOVDET"))) = 0 Then
                    mIdMovimientoDetalle = FProd.HallaMovimientoDetalle(F.NuloNumeric(.TextMatrix(A, .ColIndex("IDPARTEPRODDET"))), xCon)
                End If
                If Not FProd.GrabaCostoMovTemp(F.NuloNumeric(lblidalm.Caption), F.NuloNumeric(.TextMatrix(A, .ColIndex("IDMOVDET"))), _
                            F.NuloNumeric(.TextMatrix(A, .ColIndex("CANTIDAD"))), _
                            F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMP"))) / F.NuloNumeric(.TextMatrix(A, .ColIndex("CANTIDAD"))), _
                            F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMP"))) / F.NuloNumeric(.TextMatrix(A, .ColIndex("CANTIDAD"))), _
                            F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMP"))), _
                            F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMOD"))), _
                            F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOCIF"))), _
                            CDate(F.NuloString(.TextMatrix(A, .ColIndex("FECHA")))), _
                            "I", xCon) Then
                        
                    Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                            + vbCr + "Numero de Produccion: " & F.NuloString(.TextMatrix(A, .ColIndex("NUMPARTE"))) _
                                            + vbCr + "Item: " & F.NuloString(.TextMatrix(A, .ColIndex("ITEM"))) _
                                            + vbCr + "Cantidad: " & F.NuloNumeric(.TextMatrix(A, .ColIndex("CANTIDAD"))) _
                                            + vbCr + "Fecha de Movimiento: " & F.NuloString(.TextMatrix(A, .ColIndex("FECHA")))
                End If
                
                
                ' Actualizamos Valores
                .TextMatrix(A, .ColIndex("COSTOTOTAL")) = Format(.TextMatrix(A, .ColIndex("COSTOPRIMO")) + F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOCIF"))), FORMAT_IMPORTEKARDEX)
                .TextMatrix(A, .ColIndex("CUPROD")) = Format(.TextMatrix(A, .ColIndex("COSTOTOTAL")) / F.NuloNumeric(.TextMatrix(A, .ColIndex("CANTIDAD"))), FORMAT_IMPORTEKARDEX)
            Next A
        End With
    End If
    
    lblsumafactor.Visible = True
    lblsumafactor.Caption = "Suma MOD: " & F.NuloNumeric(GRID_SUMAR_COL(fg(0), fg(0).ColIndex("COSTOMOD"))) & ", Sum CIF: " & F.NuloNumeric(GRID_SUMAR_COL(fg(0), fg(0).ColIndex("COSTOCIF")))
    
    ' Se reprocesan los items a los cuales se ha distribuido gastos
    If mMetVal.ProcesaMP Then
        Dim mCostoTotalFaltante As Double
        Dim mContador As Long
        Set mDataBase.Connection = xCon
        mDataBase.ClearParameter
        mDataBase.CommandText = "SELECT con_librocostotemp.idmovdet, con_librocostotemp.fecha, alm_ingresodet.iditem, alm_ingreso.tipmov " _
            + vbCr + "FROM alm_ingreso INNER JOIN (alm_ingresodet INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id " _
            + vbCr + "WHERE (((con_librocostotemp.fecha)>=CDate('" & PRIMERDIAMES_ & "') And (con_librocostotemp.fecha)<=CDate('" & ULTIMODIAMES_ & "')) AND ((con_librocostotemp.idalmproc)= " & F.NuloNumeric(lblidalm.Caption) & ") AND ((alm_ingresodet.iditem) IN " _
            + vbCr + "( " _
            + vbCr + "SELECT alm_ingresodet.iditem " _
            + vbCr + "FROM alm_ingresodet INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet " _
            + vbCr + "WHERE (((con_librocostotemp.costomod)>0) AND ((con_librocostotemp.costocif)>0) AND ((con_librocostotemp.fecha)>=CDate('" & PRIMERDIAMES_ & "') And (con_librocostotemp.fecha)<=CDate('" & ULTIMODIAMES_ & "'))) " _
            + vbCr + "GROUP BY alm_ingresodet.iditem " _
            + vbCr + ") " _
            + vbCr + ")) " _
            + vbCr + "ORDER BY alm_ingresodet.iditem, con_librocostotemp.fecha, alm_ingreso.tipmov"

        Set mRecord = mDataBase.GetRecordset
        mRecord.Sort = "iditem, fecha, tipmov"
        Set mDataBase = Nothing
        If mRecord.RecordCount > 0 Then
            FWin.ShowProgress "", 0, mRecord.RecordCount
            mContador = 0
            mRecord.MoveFirst
            While Not mRecord.EOF
                mContador = mContador + 1
                FWin.SetProgress "Reprocesando movimientos con gastos distribuidos", mContador
                mCostoTotalFaltante = mCostoTotalFaltante + FProd.CalcularCostoMovimiento(F.NuloNumeric(lblidalm.Caption), F.NuloNumeric(mRecord("idmovdet")), xCon)
                If FProd.GetError Then
                    Err.Raise vbObjectError + 1
                End If
                mRecord.MoveNext
            Wend
            FWin.HideProgress
        End If
        
        '**********************************
        ' se reprocesan los items faltantes
        '**********************************
        Dim mSaldoCantidad As Double
        Dim mSaldoImporte As Double
        Dim mRecordAux As New ADODB.Recordset
        
        ' Items con Problemas
        Set mRecord = Nothing
        Set mRecord = F.GeneraRstSQL(F.SQL_MovTotalizado("", 0, PRIMERDIAMES_, ULTIMODIAMES_, xCon, True), xCon)
        If mRecord.RecordCount > 0 Then
            FWin.ShowProgress "", 0, mRecord.RecordCount
            mContador = 0
            mRecord.MoveFirst
            While Not mRecord.EOF
                mContador = mContador + 1
                FWin.SetProgress "Reprocesando Movimientos Faltantes", mContador
                mSaldoCantidad = Format(F.NuloNumeric(mRecord("canini")) + F.NuloNumeric(mRecord("canent")) - F.NuloNumeric(mRecord("cansal")), FORMAT_CANTIDAD)
                mSaldoImporte = Format(F.NuloNumeric(mRecord("costoini")) + F.NuloNumeric(mRecord("costoent")) - F.NuloNumeric(mRecord("costosal")), FORMAT_MONTO)
                If mSaldoCantidad = 0 And mSaldoImporte > 0 Then
                    ' Movimientos con Problemas
                    Dim mSQLAux As String
                    Set mRecordAux = Nothing
                    mSQLAux = "SELECT con_librocostotemp.idmovdet, con_librocostotemp.fecha, alm_ingresodet.iditem, alm_ingreso.tipmov " _
                    + vbCr + "FROM alm_ingreso INNER JOIN (alm_ingresodet INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id " _
                    + vbCr + "WHERE (((con_librocostotemp.idalmproc)= " & F.NuloNumeric(lblidalm.Caption) & ") AND ((con_librocostotemp.fecha)>=CDate('" & PRIMERDIAMES_ & "') AND (con_librocostotemp.fecha)<=CDate('" & ULTIMODIAMES_ & "')) AND ((alm_ingresodet.iditem) = " & mRecord("iditem") & " )) " _
                    + vbCr + "ORDER BY alm_ingresodet.iditem, con_librocostotemp.fecha, alm_ingreso.tipmov"
                    Set mRecordAux = F.GeneraRstSQL(mSQLAux, xCon)
                    If mRecordAux.RecordCount > 0 Then
                        mRecordAux.MoveFirst
                        While Not mRecordAux.EOF
                            mCostoTotalFaltante = mCostoTotalFaltante + FProd.CalcularCostoMovimiento(F.NuloNumeric(lblidalm.Caption), F.NuloNumeric(mRecordAux("idmovdet")), xCon)
                            If FProd.GetError Then
                                Err.Raise vbObjectError + 1
                            End If
                            mRecordAux.MoveNext
                        Wend
                    End If
                End If
                mRecord.MoveNext
            Wend
            FWin.HideProgress
        End If
    End If
    
    '******************************
    ' Si es el almacen de PT Nuevo
    '******************************
    If F.NuloNumeric(lblidalm.Caption) = 5 Then
        Set mDataBase.Connection = xCon
        mDataBase.ClearParameter
        mDataBase.CommandText = "SELECT pro_producciondet.iditem " _
            + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
            + vbCr + "WHERE (((pro_produccion.fchdoc)>=CDate('" & PRIMERDIAMES_ & "') And (pro_produccion.fchdoc)<=CDate('" & ULTIMODIAMES_ & "')) AND ((pro_produccion.idalm)=" & F.NuloNumeric(lblidalm.Caption) & ")) " _
            + vbCr + "GROUP BY pro_producciondet.iditem"
            
        Set mRecord = mDataBase.GetRecordset
        Set mDataBase = Nothing
        If mRecord.RecordCount > 0 Then
            FWin.ShowProgress "", 0, mRecord.RecordCount
            mContador = 0
            mRecord.MoveFirst
            While Not mRecord.EOF
                mContador = mContador + 1
                FWin.SetProgress "Calculando Costo Salidas", mContador
                If Not FProd.CalcularCostoSalidas(F.NuloNumeric(mRecord("iditem")), _
                                    F.NuloNumeric(lblidalm.Caption), _
                                    PRIMERDIAMES_, ULTIMODIAMES_, xCon) Then
                        
                    Err.Raise &HFFFFFF01, , "Error interno al intentar procesar el costo de las salidas."
                End If
                mRecord.MoveNext
            Wend
            FWin.HideProgress
        End If
        
        '*************************
        ' Procesar Faltantes
        '*************************
        If ckProcFal.Value Then
'            Set mDataBase.Connection = xCon
'            mDataBase.ClearParameter
'            mDataBase.CommandText = "SELECT alm_ingreso.id AS idmov, alm_ingresodet.idmovdet, alm_ingreso.idtipdocref, alm_ingreso.iddocref, alm_ingreso.idtipdocref2, alm_ingreso.iddocref2, alm_ingreso.fching, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS numdoc, alm_ingresodet.iditem, [con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad] AS costo, alm_inventario.codpro AS coditem, alm_inventario.descripcion AS item, alm_ingresodet.cantidad " _
'                    + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
'                    + vbCr + "WHERE (((alm_ingresodet.cantidad) > 0) AND (([con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad]) Is Null) AND ((alm_ingreso.fching)>=CDate('" & mFechaInicioInventario & "')) AND ((alm_ingreso.ano)=" & AnoTra & ") AND ((alm_ingreso.idmes)=" & MESACTUAL_ & "))" _
'                    + vbCr + "ORDER BY alm_ingreso.fching, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc], alm_ingresodet.iditem"
'            Set mRecord = mDataBase.GetRecordset
'            Set mDataBase = Nothing
'            If mRecord.RecordCount > 0 Then
'                FWin.ShowProgress "Proc.Mov.Fal.: ", 0, mRecord.RecordCount
'                mContador = 0
'                mRecord.MoveFirst
'                While Not mRecord.EOF
'                    mContador = mContador + 1
'                    FWin.SetProgress F.NuloString(mRecord("fching") & " - " & mRecord("numdoc") & " - " & mRecord("item") & " - " & mRecord("cantidad")), mContador
'                    mCostoTotalFaltante = mCostoTotalFaltante + FProd.CalcularCostoMovimiento(F.NuloNumeric(lblidalm.Caption), F.NuloNumeric(mRecord("idmovdet")), xCon)
'                    If FProd.GetError Then
'                        Err.Raise vbObjectError + 1
'                    End If
'                    mRecord.MoveNext
'                Wend
'                FWin.HideProgress
'            End If
        End If
    End If
    
    FraProgreso.Visible = False
    Agregando = False
    Exit Sub
    
ERROR_:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    FWin.HideProgress
    Agregando = False
    MsgBox "No se pudo procesar los gastos indirectos por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub pLlenarDatos(MESATRABAJAR_ As Integer)
    Dim Rpta As Integer
    Dim F As New SistemaLogica.Funciones
    Dim FWin As New SistemaWindows.SistemaWindowsClass
    Dim mMetVal As New ContabilidadEntidad.EConfigVal
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mLParteProdDet As New ProduccionEntidad.LEParteProdDet
    Dim FProd As New ProduccionLogica.Funciones
    Dim mFechaInicioInventario As Date
    Dim ANIOACTUAL_ As Integer
    Dim MESACTUAL_ As Integer
    Dim PRIMERDIAMES_ As Date
    Dim ULTIMODIAMES_ As Date
    Dim mCantidad As Double
    Dim LparteProdDet As New ProduccionEntidad.LEParteProdDet
    Dim A As Integer
    Dim mIndiceTotal As Double
    Dim mIndice As Double
    
On Error GoTo BloqueError

    ' Se valida la fecha de cierre de mes
    If F.MesCerradoOpcion(MESATRABAJAR_, CLng(F.KeyValue("IdOpcionSistemaMovimientoAlmacen", xCon)), xCon) Then
        Rpta = MsgBox("El presente mes para la opcion: ingresos y salidas de almacén, no se encuentra cerrado ¿ Esta seguro desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1, "Cierre de Mes")
        If Rpta = vbNo Then
            Exit Sub
        End If
    End If
    ' Se validan datos
    If F.NuloNumeric(lblidmetval.Caption) = 0 Then
        F.MostrarMensajeError "Debe seleccionar un metodo de valorización", "Error"
        Exit Sub
    End If
    ' Se carga la configuracion de la valorizacion
    Set mMetVal.Conexion = xCon
    mMetVal.Fetch NulosN(lblidmetval)

    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    mFechaInicioInventario = F.FechaInicioMovimientos(F.NuloNumeric(lblidalm.Caption), xCon) 'CDate("01/05/2014")
    Agregando = True
    
    Me.MousePointer = vbHourglass
    
    ' Se reprocesa todo
    If ReprocesarCheck.Value = 1 Then
        Set mDataBase.Connection = xCon
        ' Se actualiza cantidades de Kardex
        mDataBase.ClearParameter
        mDataBase.CommandText = "DELETE FROM con_librocostotemp " _
            + vbCr + "WHERE (((con_librocostotemp.idalmproc = " & F.NuloNumeric(lblidalm.Caption) & ")) AND ((con_librocostotemp.fecha)>=CDate('" & PRIMERDIAMES_ & "') AND (con_librocostotemp.fecha)<=CDate('" & ULTIMODIAMES_ & "')))"
        mDataBase.Execute
    End If
        
    If mMetVal.ProcesaMP Then
        ' Se generan los Costos Faltantes
        Dim mCostoTotalFaltante As Double
        Dim mContador As Long
        Set mDataBase.Connection = xCon
        mDataBase.ClearParameter
        mDataBase.CommandText = "SELECT alm_ingreso.id AS idmov, alm_ingresodet.idmovdet, alm_ingreso.idtipdocref, alm_ingreso.iddocref, alm_ingreso.idtipdocref2, alm_ingreso.iddocref2, alm_ingreso.fching, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS numdoc, alm_ingresodet.iditem, [con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad] AS costo, alm_inventario.codpro AS coditem, alm_inventario.descripcion AS item, alm_ingresodet.cantidad " _
                + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
                + vbCr + "WHERE (((alm_ingreso.idalm) <> 5) AND ((alm_ingresodet.cantidad) > 0) AND (([con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad]) Is Null) AND ((alm_ingreso.fching)>=CDate('" & mFechaInicioInventario & "')) AND ((alm_ingreso.ano)=" & AnoTra & ") AND ((alm_ingreso.idmes)=" & MESACTUAL_ & ")) " _
                + vbCr + "ORDER BY alm_ingreso.fching, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc], alm_ingresodet.iditem"
        Set mRecord = mDataBase.GetRecordset
        Set mDataBase = Nothing
        If mRecord.RecordCount > 0 Then
            FWin.ShowProgress "Proc.Mov.: ", 0, mRecord.RecordCount
            mContador = 0
            mRecord.MoveFirst
            While Not mRecord.EOF
                mContador = mContador + 1
                FWin.SetProgress F.NuloString(mRecord("fching") & " - " & mRecord("numdoc") & " - " & mRecord("item") & " - " & mRecord("cantidad")), mContador
                mCostoTotalFaltante = mCostoTotalFaltante + FProd.CalcularCostoMovimiento(F.NuloNumeric(lblidalm.Caption), F.NuloNumeric(mRecord("idmovdet")), xCon)
                If FProd.GetError Then
                    Err.Raise vbObjectError + 1
                End If
                mRecord.MoveNext
            Wend
            FWin.HideProgress
        End If

'        ' MODIFICADO 20150913 PARA REVISAR EL NUEVO PROCESO SOLICITADO
'        ' SE REGRESA A LA FUNCIONALIDAD ANTERIOR
'        mCostoTotalFaltante = FProd.CostoPrimo(F.NuloNumeric(lblidalm.Caption), PRIMERDIAMES_, ULTIMODIAMES_, mFechaInicioInventario, xCon)
    End If
    ' Se llenan los datos
    Set LparteProdDet.Conexion = xCon
    LparteProdDet.LoadChild = False
    LparteProdDet.Fetch , F.NuloNumeric(lblidalm.Caption), PRIMERDIAMES_, ULTIMODIAMES_, mFechaInicioInventario
    ' Se recorre la lista para calcular su importe unitario
    With fg(0)
        .Rows = .FixedRows
        Dim ParteProdDet As New ProduccionEntidad.EParteProdDet
        For Each ParteProdDet In LparteProdDet
            Dim mImporteMP As Double
            Dim mTipo As Integer
            
            mImporteMP = F.NuloNumeric(F.BuscaCodigoTabla(ParteProdDet.IdMovimientoDetalle, "idmovdet", "costoprimo", "con_librocostotemp", "N", xCon))
            mTipo = F.NuloNumeric(F.BuscaCodigoTabla(ParteProdDet.IdItem, "id", "tippro", "alm_inventario", "N", xCon))
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("FECHA")) = F.NuloString(ParteProdDet.Fecha)
            .TextMatrix(.Rows - 1, .ColIndex("NUMPARTE")) = F.NuloString(ParteProdDet.OrdenProduccion)
            If mTipo = 3 Then
                .TextMatrix(.Rows - 1, .ColIndex("TIPO")) = "PT"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("TIPO")) = "PI"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = F.NuloString(ParteProdDet.Item)
            .TextMatrix(.Rows - 1, .ColIndex("RECETA")) = F.NuloString(ParteProdDet.Receta)
            .TextMatrix(.Rows - 1, .ColIndex("UM")) = F.NuloString(ParteProdDet.UnidadMedida)
            .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = F.NuloString(ParteProdDet.CantidadProducida)
            .TextMatrix(.Rows - 1, .ColIndex("HORINI")) = Format(F.NuloString(ParteProdDet.HoraInicio), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, .ColIndex("HORFIN")) = Format(F.NuloString(ParteProdDet.HoraFin), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, .ColIndex("THORAS")) = Format(CDate(ParteProdDet.HoraFin) - CDate(ParteProdDet.HoraInicio), "HH:mm")
            .TextMatrix(.Rows - 1, .ColIndex("THORASNUM")) = F.ConvertirHoraADecimal(CDate(ParteProdDet.HoraFin) - CDate(ParteProdDet.HoraInicio))
            .TextMatrix(.Rows - 1, .ColIndex("COSTOMP")) = Format(mImporteMP, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("IDITEM")) = ParteProdDet.IdItem
            .TextMatrix(.Rows - 1, .ColIndex("IDPARTEPRODDET")) = ParteProdDet.IdParteProduccionDet
            .TextMatrix(.Rows - 1, .ColIndex("IDMOVDET")) = ParteProdDet.IdMovimientoDetalle
        Next
        '*************************************
        ' Factor de Distribucion
        '*************************************
        Select Case mMetVal.CodigoTipoItemDistribucion
            Case "PT"
                ' Se Suma la cantidad de solo PT
                For A = fg(0).FixedRows To fg(0).Rows - 1
                    If F.NuloString(fg(0).TextMatrix(A, fg(0).ColIndex("TIPO"))) = "PT" Then
                        mIndiceTotal = mIndiceTotal + F.NuloNumeric(fg(0).TextMatrix(A, fg(0).ColIndex(mMetVal.ColumnaFactorDistribucion)))
                    End If
                Next
                
            Case "TD"
                mIndiceTotal = F.NuloNumeric(GRID_SUMAR_COL(fg(0), fg(0).ColIndex(mMetVal.ColumnaFactorDistribucion)))
                
            Case Else ' Manual
                mIndiceTotal = fg(0).Rows - 1
                
        End Select
        For A = .FixedRows To .Rows - 1
            Select Case mMetVal.CodigoTipoItemDistribucion
                Case "PT"
                    mIndice = F.NuloNumeric(.TextMatrix(A, .ColIndex(mMetVal.ColumnaFactorDistribucion)))
                    If F.NuloString(.TextMatrix(A, .ColIndex("TIPO"))) = "PT" Then
                        .TextMatrix(A, .ColIndex("FACTOR")) = Format(((mIndice / mIndiceTotal) * 100), FORMAT_IMPORTEKARDEX)
                    Else
                        .TextMatrix(A, .ColIndex("FACTOR")) = Format(0, FORMAT_CANTIDAD)
                    End If
                    
                Case "TD"
                    mIndice = F.NuloNumeric(.TextMatrix(A, .ColIndex(mMetVal.ColumnaFactorDistribucion)))
                    If mIndiceTotal = 0 Then
                        .TextMatrix(A, .ColIndex("FACTOR")) = Format(0, FORMAT_IMPORTEKARDEX)
                    Else
                        .TextMatrix(A, .ColIndex("FACTOR")) = Format(((mIndice / mIndiceTotal) * 100), FORMAT_IMPORTEKARDEX)
                    End If
                    
                Case Else
                    mIndice = 1
                    .TextMatrix(A, .ColIndex("COSTOMOD")) = Format(((mIndice / mIndiceTotal) * 100), FORMAT_CANTIDAD)
            End Select
        Next A
    End With
    
    Me.MousePointer = vbDefault
    Agregando = False
    Exit Sub
    
BloqueError:
    Agregando = False
    FWin.HideProgress
    Me.MousePointer = vbDefault
    F.MostrarMensajeError Err.Description, "LlenarDatos", Err.Source, Err.Number
End Sub

Private Sub llenarDetalleInsumos()
    If Agregando Then Exit Sub
    
    fg(3).Rows = fg(3).FixedRows
    With fg(3)
        ' Obtenemos el Parte detallado seleccionado
        Dim mParteProdDet As New ProduccionEntidad.EParteProdDet
        mParteProdDet.LoadChild = True
        Set mParteProdDet.Conexion = xCon
        mParteProdDet.Fetch F.NuloNumeric(fg(0).TextMatrix(fg(0).Row, fg(0).ColIndex("IDPARTEPRODDET")))
        
        Dim mParteProdDetIns As New ProduccionEntidad.EParteProdDetIns
        For Each mParteProdDetIns In mParteProdDet.LParteProduccionDetIns
            Dim mParteProdDetInsMov As New ProduccionEntidad.EParteProdDetInsMov
            For Each mParteProdDetInsMov In mParteProdDetIns.LParteProduccionDetInsMov
                Dim mIdMovimiento As Integer
                Dim mTipoMovimiento As Integer
                Dim mCostoUnitarioPromedio As Double
                Dim mImporte As Double
                
                mIdMovimiento = F.NuloNumeric(F.BuscaCodigoTabla(mParteProdDetInsMov.IdMovimientoDetalle, "idmovdet", "id", "alm_ingresodet", "N", xCon))
                mTipoMovimiento = F.NuloNumeric(F.BuscaCodigoTabla(mIdMovimiento, "id", "tipmov", "alm_ingreso", "N", xCon))
                mCostoUnitarioPromedio = F.NuloNumeric(F.BuscaCodigoTabla(mParteProdDetInsMov.IdMovimientoDetalle, "idmovdet", "costounitariopromedio", "con_librocostotemp", "N", xCon))
                mImporte = F.NuloNumeric(F.BuscaCodigoTabla(mParteProdDetInsMov.IdMovimientoDetalle, "idmovdet", "costoprimo", "con_librocostotemp", "N", xCon))
                
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("CODIGOITEM")) = mParteProdDetIns.CodigoItem
                .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = mParteProdDetIns.Item
                If mTipoMovimiento = 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("TIPOMOV")) = "S"
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("TIPOMOV")) = "I"
                End If
                .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = mParteProdDetInsMov.Cantidad
                .TextMatrix(.Rows - 1, .ColIndex("COSTOPROMEDIO")) = Format(mCostoUnitarioPromedio, FORMAT_IMPORTEKARDEX)
                .TextMatrix(.Rows - 1, .ColIndex("IMPORTE")) = Format(mImporte, FORMAT_IMPORTEKARDEX)
            Next
        Next
        ' Se agrega la fila de totales
        .Rows = .Rows + 1
        FORMATO_CELDA fg(3), .Rows - 1, .ColIndex("COSTOPROMEDIO"), , True, , "TOTAL"
        .TextMatrix(.Rows - 1, .ColIndex("IMPORTE")) = Format(GRID_SUMAR_COL(fg(3), .ColIndex("IMPORTE")), FORMAT_MONTO)
        .TopRow = .Rows - 1
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim MENSAJE_ As String
    Dim nSQLId As String
    Dim nSQLId2 As String
    Dim NUMEROMAXTRAB_ As Integer
    Dim NUMREGAAGREGAR_ As Integer
    
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Integer
    Dim MESACTUAL_ As Integer
        
    If QueHace = 3 Then Exit Sub

    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = cbMes.ListIndex + 1
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESACTUAL_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = cbMes.ListIndex + 1
            
    Select Case Index
        Case 0 ' METODO DE VALORIZACION
            ReDim xCampos(2, 4) As String

            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "abrev":            xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                        
            cSQL = "SELECT * FROM mae_configval;"
                
            nTitulo = "Buscando Metodos de valorizacion"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblidmetval.Caption = F.NuloNumeric(xRs("id"))
            txtmetval.Text = F.NuloString(xRs("abrev"))
            lblmetval.Caption = F.NuloString(xRs("descripcion"))
            cmd(2).SetFocus
            
        Case 2 ' CONSULTAR
            pLlenarDatos MESACTUAL_
            
        Case 3 ' PROCESAR
            pProcesarDatos MESACTUAL_
            
        Case 4 ' Almacen
            ReDim xCampos(2, 4) As String

            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "codigo":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                        
            cSQL = "SELECT * FROM alm_almacenes WHERE idtipalm = 1"
                
            nTitulo = "Buscando Metodos de valorizacion"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblidalm.Caption = F.NuloNumeric(xRs("id"))
            txtalm.Text = F.NuloString(xRs("codigo"))
            lblalm.Caption = F.NuloString(xRs("descripcion"))
            ' Metodo de valorizacion por defecto
            lblidmetval.Caption = F.NuloNumeric(xRs("idmetval"))
            lblmetval.Caption = F.NuloString(F.BuscaCodigoTabla(F.NuloNumeric(xRs("idmetval")), "id", "descripcion", "mae_configval", "N", xCon))
            txtmetval.Text = F.NuloString(F.BuscaCodigoTabla(F.NuloNumeric(xRs("idmetval")), "id", "abrev", "mae_configval", "N", xCon))
            
            txtmetval.SetFocus
            
    End Select
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, F.NuloNumeric(RstLibro("id")), xCon
    End If
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLibro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLibro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
        Case 0
            Select Case Col
                Case fg(0).ColIndex("FACTOR")
                    lblsumafactor.Visible = True
                    lblsumafactor.Caption = "Suma Factor: " & Format(F.NuloNumeric(GRID_SUMAR_COL(fg(0), fg(0).ColIndex("FACTOR"))), FORMAT_IMPORTEKARDEX)
            End Select
    End Select
End Sub

Private Sub fg_DblClick(Index As Integer)
    If Index <> 0 Then Exit Sub
    If Agregando Then Exit Sub
    If (fg(0).Row = fg(0).Rows - 1) And QueHace = 3 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    llenarDetalleInsumos
    Me.MousePointer = vbDefault
End Sub

Private Sub cbMes_DropDown()
    If Agregando Then Exit Sub
End Sub

Private Sub fg_EnterCell(Index As Integer)
    Select Case Index
        Case 0
            If QueHace = 3 Then fg(0).Editable = flexEDNone: Exit Sub
            Select Case fg(0).Col
                Case fg(0).ColIndex("FACTOR")
                    fg(0).Editable = flexEDKbdMouse

                Case Else
                    fg(0).Editable = flexEDNone
            End Select

        Case Else
            fg(Index).Editable = flexEDNone
    End Select
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Index
        Case 0
            Select Case Col
                Case fg(0).ColIndex("FACTOR")
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
            End Select
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        mMesActivo = xMes
            
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        pCargarGrid
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        '--interrumpir
        'BANDERA_ = True
    End If
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 750
    
    Label4(0).Width = Me.Width - 100
    Dg1.Width = TabOne1.Width - 100
    Dg1.Height = TabOne1.Height - 1000
    
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    
    Frame4.Width = TabOne1.Width - 150
    Frame4.Height = TabOne1.Height - 2655
    
    fg(0).Width = Frame4.Width - 150
    fg(0).Height = Frame4.Height - 3480
    
    Frame7.Top = Frame4.Height - 3135
    Frame7.Width = Frame4.Width - 210
    
    fg(3).Width = Frame7.Width - 195
End Sub

Private Sub iniciarCampos()
    TabOne1.CurrTab = 0
        
    '**********************
    ' CONFIGURACIONES GRID
    '**********************
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).Editable = flexEDKbdMouse
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    fg(0).Rows = fg(0).FixedRows
    fg(0).FrozenCols = fg(0).ColIndex("ITEM")
    
    
    fg(3).AllowUserResizing = flexResizeColumns
    fg(3).AutoSearch = flexSearchFromTop
    fg(3).ExplorerBar = flexExSortShow
    fg(3).SelectionMode = flexSelectionByRow
    fg(3).ForeColorSel = &H80000005
    fg(3).BackColorSel = &H80&
    fg(3).Editable = flexEDKbdMouse
    fg(3).Rows = fg(3).FixedRows
    
    Llenar_Mes cbMes
    
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub pCargarGrid()
    Dim cSQL  As String
    Dim Rpta As Integer
    
    TDB_FiltroLimpiar Dg1
    
    cSQL = "SELECT con_librocosto.*, con_meses.descripcion AS desmes, mae_metodoval.descripcion AS desmetval, alm_almacenes.descripcion AS desalm " _
        + vbCr + "FROM ((con_librocosto LEFT JOIN mae_metodoval ON con_librocosto.idmetodoval = mae_metodoval.id) LEFT JOIN con_meses ON con_librocosto.idmes = con_meses.id) LEFT JOIN alm_almacenes ON con_librocosto.idalm = alm_almacenes.id " _
        + vbCr + "ORDER BY con_librocosto.idmes"
        
    Me.MousePointer = vbHourglass
    
    RST_Busq RstLibro, cSQL, xCon
    Set Dg1.DataSource = RstLibro
    
    Me.MousePointer = vbDefault
    If RstLibro.State = 0 Then Exit Sub
End Sub

Private Sub MuestraSegundoTab()
    Dim mLibroCosto As New ContabilidadEntidad.ELibroCosto
    Dim IMPORTEPRODUCCION_ As Double
    
On Error GoTo BloqueError
        
    If RstLibro.RecordCount = 0 Then Exit Sub
    If RstLibro.EOF = True Then Exit Sub
    
    Blanquea
    Me.MousePointer = vbHourglass
    Agregando = True
    Set mLibroCosto.Conexion = xCon
    mLibroCosto.LoadChild = True
    mLibroCosto.Fetch F.NuloNumeric(RstLibro("id"))
     
    cbMes.ListIndex = mLibroCosto.IdMes - 1
    txtdescripcion.Text = mLibroCosto.Descripcion
    lblidmetval.Caption = mLibroCosto.IdMetodoVal
    txtmetval.Text = mLibroCosto.CodigoMetodoVal
    lblmetval.Caption = mLibroCosto.MetodoVal
    lblidalm.Caption = mLibroCosto.IdAlmacen
    txtalm.Text = mLibroCosto.CodigoAlmacen
    lblalm.Caption = mLibroCosto.Almacen
    MODText.Text = mLibroCosto.CostoMOD
    CIFText.Text = mLibroCosto.CostoCIF
    
    Dim mLibroCostoDet As ContabilidadEntidad.ELibroCostoDet
    fg(0).Rows = fg(0).FixedRows
    For Each mLibroCostoDet In mLibroCosto.LLibroCostoDet
        With fg(0)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("FECHA")) = Format(mLibroCostoDet.Fecha, FORMAT_DATE)
            .TextMatrix(.Rows - 1, .ColIndex("NUMPARTE")) = mLibroCostoDet.ParteProd
            .TextMatrix(.Rows - 1, .ColIndex("TIPO")) = mLibroCostoDet.Tipo
            .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = mLibroCostoDet.Item
            .TextMatrix(.Rows - 1, .ColIndex("RECETA")) = mLibroCostoDet.Receta
            .TextMatrix(.Rows - 1, .ColIndex("UM")) = mLibroCostoDet.UniMed
            .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = Format(mLibroCostoDet.Cantidad, FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, .ColIndex("HORINI")) = Format(mLibroCostoDet.HoraInicio, FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, .ColIndex("HORFIN")) = Format(mLibroCostoDet.HoraFin, FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, .ColIndex("THORAS")) = Format(mLibroCostoDet.TotalHoras, "HH:mm")
            .TextMatrix(.Rows - 1, .ColIndex("THORASNUM")) = F.ConvertirHoraADecimal(mLibroCostoDet.TotalHoras)
            .TextMatrix(.Rows - 1, .ColIndex("FACTOR")) = Format(mLibroCostoDet.FactorDistribucion, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOMP")) = Format(mLibroCostoDet.ImporteMP, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOMOD")) = Format(mLibroCostoDet.ImporteMOD, FORMAT_MONTO)
            IMPORTEPRODUCCION_ = mLibroCostoDet.ImporteMP + mLibroCostoDet.ImporteMOD
            .TextMatrix(.Rows - 1, .ColIndex("COSTOPRIMO")) = Format(IMPORTEPRODUCCION_, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOCIF")) = Format(mLibroCostoDet.ImporteCIF, FORMAT_IMPORTEKARDEX)
            IMPORTEPRODUCCION_ = IMPORTEPRODUCCION_ + mLibroCostoDet.ImporteCIF
            .TextMatrix(.Rows - 1, .ColIndex("COSTOTOTAL")) = Format(IMPORTEPRODUCCION_, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, .ColIndex("CUPROD")) = Format(IMPORTEPRODUCCION_ / mLibroCostoDet.Cantidad, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, .ColIndex("IDLIBRODET")) = mLibroCostoDet.IdLibroCostoDet
            .TextMatrix(.Rows - 1, .ColIndex("IDPARTEPROD")) = mLibroCostoDet.IdParteProd
            .TextMatrix(.Rows - 1, .ColIndex("IDPARTEPRODDET")) = mLibroCostoDet.IdParteDetalle
            .TextMatrix(.Rows - 1, .ColIndex("IDMOVDET")) = mLibroCostoDet.IdMovimientoDetalle
            .TextMatrix(.Rows - 1, .ColIndex("IDITEM")) = mLibroCostoDet.IdItem
        End With
    Next
    With fg(0)
        ' Se agrega la fila de totales
        .Rows = .Rows + 1
        FORMATO_CELDA fg(0), .Rows - 1, .ColIndex("HORFIN"), , True, , "TOTAL"
        .TextMatrix(.Rows - 1, .ColIndex("COSTOMP")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOMP")), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOMOD")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOMOD")), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOPRIMO")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOPRIMO")), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOCIF")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOCIF")), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOTOTAL")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOTOTAL")), FORMAT_MONTO)
        .TopRow = .Rows - 1
    End With
    Me.MousePointer = vbDefault
    Agregando = False
    Exit Sub
    
BloqueError:
    Me.MousePointer = vbDefault
    Agregando = False
    F.MostrarMensajeError "Ocurrio un error al cargar el detalle: " & Err.Description, "Error"
End Sub

Sub Cancelar()
    Bloquea
    Label5.Caption = "Detalle de Libro de Costos"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
     
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).Editable = flexEDNone
    fg(3).SelectionMode = flexSelectionByRow
    fg(3).Editable = flexEDNone
    
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    Bloquea
    Blanquea
    fg(0).Rows = fg(0).FixedRows
    'fg(1).Rows = fg(1).FixedRows
    'fg(2).Rows = fg(2).FixedRows
    fg(3).Rows = fg(3).FixedRows
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Libro de Costo de Producción"
    ' Grid principal
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    ' Grid auxiliar
    fg(3).Editable = flexEDKbdMouse
    fg(3).SelectionMode = flexSelectionFree
End Sub

Sub Bloquea()
    cbMes.Locked = Not cbMes.Locked
    txtdescripcion.Locked = Not txtdescripcion.Locked
    txtmetval.Locked = Not txtmetval.Locked
    txtalm.Locked = Not txtalm.Locked
    MODText.Locked = Not MODText.Locked
    CIFText.Locked = Not CIFText.Locked
    habilitar cmd, Not cmd(0).Enabled
End Sub

Sub Blanquea()
    txtdescripcion.Text = ""
    txtmetval.Text = ""
    txtalm.Text = ""
    lblmetval.Caption = ""
    lblalm.Caption = ""
    lblidalm.Caption = 0
    lblidmetval.Caption = 0
    MODText.Text = ""
    CIFText.Text = ""
    fg(3).Rows = fg(3).FixedRows
    lblsumafactor.Caption = ""
End Sub

Function Grabar() As Boolean
    Dim mIdAplicaDistribucion As Integer
    Dim mIdTipoDistribucion As Integer
    Dim mIdCampoDistribucion As Integer
    Dim A As Integer
    Dim mLibroCosto As New ContabilidadEntidad.ELibroCosto
    
    If txtdescripcion.Text = "" Then
        MsgBox "No ha especificado una descripcion para el libro actual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtdescripcion.SetFocus
        Exit Function
    End If
    
    If txtmetval.Text = "" Then
        MsgBox "No ha especificado el metodo de valorización", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtmetval.SetFocus
        Exit Function
    End If
    
    If txtalm.Text = "" Then
        MsgBox "No ha especificado el almacen de valorización", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtalm.SetFocus
        Exit Function
    End If
    
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No se han procesado datos de producción para el libro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
      
On Error GoTo ERROR_
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "modificar") + " el Libro de Costo", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Function
    
    Me.MousePointer = vbHourglass
    ' Se valida la ultima fila de totales para eliminarla
    If F.NuloString(fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("HORFIN"))) = "TOTAL" Then
        fg(0).Rows = fg(0).Rows - 1
    End If
    ' Se cargar los valores
    If QueHace = 1 Then mLibroCosto.IdLibroCosto = 0 Else mLibroCosto.IdLibroCosto = F.NuloNumeric(RstLibro("id"))
    mLibroCosto.IdMes = cbMes.ListIndex + 1
    mLibroCosto.IdMetodoVal = F.NuloNumeric(lblidmetval.Caption)
    mLibroCosto.Descripcion = F.NuloString(txtdescripcion.Text)
    mLibroCosto.IdAlmacen = NulosN(lblidalm.Caption)
    mLibroCosto.CostoMOD = F.NuloNumeric(MODText.Text)
    mLibroCosto.CostoCIF = F.NuloNumeric(CIFText.Text)
    With fg(0)
        For A = .FixedRows To .Rows - 1
            Dim mLibroCostoDet As New ContabilidadEntidad.ELibroCostoDet
            
            mLibroCostoDet.IdItem = F.NuloNumeric(.TextMatrix(A, .ColIndex("IDITEM")))
            mLibroCostoDet.IdParteDetalle = F.NuloNumeric(.TextMatrix(A, .ColIndex("IDPARTEPRODDET")))
            mLibroCostoDet.IdMovimientoDetalle = F.NuloNumeric(.TextMatrix(A, .ColIndex("IDMOVDET")))
            mLibroCostoDet.Cantidad = F.NuloNumeric(.TextMatrix(A, .ColIndex("CANTIDAD")))
            mLibroCostoDet.ImporteMP = F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMP")))
            mLibroCostoDet.ImporteMOD = F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOMOD")))
            mLibroCostoDet.ImporteCIF = F.NuloNumeric(.TextMatrix(A, .ColIndex("COSTOCIF")))
            mLibroCostoDet.FactorDistribucion = F.NuloNumeric(.TextMatrix(A, .ColIndex("FACTOR")))
            mLibroCostoDet.Tipo = F.NuloString(.TextMatrix(A, .ColIndex("TIPO")))
            ' Se agrega al padre
            mLibroCosto.LLibroCostoDet.Add mLibroCostoDet
            Set mLibroCostoDet = Nothing
        Next
    End With
    
    ' Se graba el registro
    Set mLibroCosto.Conexion = xCon
    If Not mLibroCosto.Save(CLng(xIdUsuario), F.MachineName) Then Err.Raise &HFFFFFF01, , "No se puedo guardar el registro"
    Me.MousePointer = vbDefault
    MsgBox "El registro se grabó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    mIdRegistro = mLibroCosto.IdLibroCosto
    Set mLibroCosto = Nothing
    Grabar = True
    Exit Function
    
ERROR_:
    Me.MousePointer = vbDefault
    Set mLibroCosto = Nothing
    Grabar = False
    MsgBox "No se pudo grabar el registro por el siguiente motivo :" + Trim(Err.Description)
End Function

Sub Modificar()
    If RstLibro.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, xTitulo
        Exit Sub
    End If
   
    QueHace = 2
    xHorIni = Time
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Modificando Libro de Costo de Producción"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    fg(3).Editable = flexEDKbdMouse
    fg(3).SelectionMode = flexSelectionFree
    
    xHorIni = Time
    cbMes.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs As New ADODB.Recordset

On Error GoTo BloqueError
    If RstLibro.RecordCount = 0 Then
        MsgBox "No hay registros para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar el registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)

    If Rpta = vbYes Then
        Dim mLibroCosto As New ContabilidadEntidad.ELibroCosto
                
        Set mLibroCosto.Conexion = xCon
        mLibroCosto.IdLibroCosto = F.NuloNumeric(RstLibro("id"))
        mLibroCosto.Delete CLng(xIdUsuario), F.MachineName
        
        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstLibro.Requery
        Dg1.Refresh
    End If
    Exit Sub
    
BloqueError:
    F.MostrarMensajeError "No se pudo eliminar el registro por el siguiente motivo :" + Trim(Err.Description), "Grabar"
    Set mLibroCosto = Nothing
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstLibro.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstLibro.RecordCount = 0 Then
            MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
            Exit Sub
        End If
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLibro.Requery
            Dg1.Refresh
            If RstLibro.RecordCount <> 0 Then
                RstLibro.MoveFirst
                RstLibro.Find "id=" & mIdRegistro
                If RstLibro.EOF = True Then RstLibro.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstLibro.Filter = "": TDB_FiltroLimpiar Dg1
    End If
        
    If Button.Index = 14 Then ExportarExcel fg(0)
    
    If Button.Index = 17 Then Unload Me
End Sub

Sub ExportarExcel(ByRef GRID_ As VSFlexGrid)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE PRODUCCIÓN"
    
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub
