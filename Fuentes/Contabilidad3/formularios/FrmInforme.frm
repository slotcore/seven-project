VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Diseño de Informes"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11775
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
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":2706
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":2B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":2FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInforme.frx":32C4
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
      Width           =   11775
      _ExtentX        =   20770
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
            Object.ToolTipText     =   "Agregar"
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
      TabIndex        =   1
      Top             =   360
      Width           =   11745
      _cx             =   20717
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
         Height          =   6795
         Left            =   45
         TabIndex        =   4
         Top             =   375
         Width           =   11655
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
            Columns(2).Caption=   "Cant. Fila"
            Columns(2).DataField=   "canfila"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=7117"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7038"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1931"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1852"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
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
            TabIndex        =   6
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   12390
         TabIndex        =   2
         Top             =   375
         Width           =   11655
         Begin VB.TextBox TxtDescripcion 
            Height          =   315
            Left            =   1335
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   11
            Text            =   "TxtDescripcion"
            Top             =   495
            Width           =   6720
         End
         Begin VB.TextBox TxtColumna 
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "null"
            Text            =   "TxtColumna"
            Top             =   870
            Width           =   1215
         End
         Begin VB.Frame fra 
            Height          =   645
            Index           =   1
            Left            =   8160
            TabIndex        =   7
            Top             =   540
            Width           =   3210
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar"
               Height          =   375
               Index           =   0
               Left            =   150
               TabIndex        =   9
               ToolTipText     =   "Agregar Cuenta Contable"
               Top             =   180
               Width           =   1260
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Eliminar"
               Height          =   375
               Index           =   1
               Left            =   1620
               TabIndex        =   8
               ToolTipText     =   "Eliminar Cuenta"
               Top             =   180
               Width           =   1260
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5445
            Left            =   60
            TabIndex        =   12
            Top             =   1260
            Width           =   11325
            _cx             =   19976
            _cy             =   9604
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmInforme.frx":3656
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Top             =   630
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. Columnas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   13
            ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Informe"
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
Attribute VB_Name = "FrmInforme"
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
Dim mIdRegistro& '--identificador del registro

Dim xHorIni As Date
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Sub Cancelar()
    QueHace = 3
    Bloquea False
    ActivaTool
    Label5.Caption = "Detalle del Informe"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
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
    If ColIndex = 5 Then Exit Sub
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


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Categoria":    xCampos(1, 1) = "cat":              xCampos(1, 2) = "1200":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Variable":     xCampos(2, 1) = "variable":         xCampos(2, 2) = "800":     xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fórmula":      xCampos(3, 1) = "formula":          xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    
    
    '*******************************************************************************************
    Dim nSQLId As String
'''    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "alm_inventario.id", " NOT IN ", True)
'''    If nSQLId <> "" Then nSQLId = " AND " & nSQLId
    '*******************************************************************************************
    '--obs. apareceran solo items de ventas que tengan cuenta contable
    
    xform.SQLCad = "SELECT con_concepto.id, con_conceptocat.descripcion AS cat, con_concepto.descripcion, con_concepto.variable, con_concepto.formula " _
                & " FROM con_concepto INNER JOIN con_conceptocat ON con_concepto.idcat = con_conceptocat.id WHERE con_concepto.activo =-1"

    
    xform.Titulo = "Buscando Concepto"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    
    xform.FormaBusca = CualquierParte
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Dim A As Integer
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            fg1.TextMatrix(fg1.Row, Col - 1) = NulosC(xRs("id"))
            fg1.TextMatrix(fg1.Row, Col) = NulosC(xRs("descripcion"))
            fg1.TextMatrix(fg1.Row, Col + 1) = NulosC(xRs("variable"))
        End If
    End If
    '------------
    '------------
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
Dim A As Integer
    If Row < 1 Then Exit Sub
    '--recorrer todas las columnas
    For A = 0 To NulosN(TxtColumna.Text) - 1
        
        
'        If Col < 4 + (4 * A) Then Exit For
        
        If Col = 4 + (4 * A) Then
            If NulosN(fg1.TextMatrix(Row, Col)) = 0 Then
                FORMATO_CELDA fg1, Row, Col - 1, vbBlack, False
                FORMATO_CELDA fg1, Row, Col - 2, vbBlack, False
                FORMATO_CELDA fg1, Row, Col - 3, vbBlack, False
            Else
                FORMATO_CELDA fg1, Row, Col - 1, vbBlack, True
                FORMATO_CELDA fg1, Row, Col - 2, vbBlack, True
                FORMATO_CELDA fg1, Row, Col - 3, vbBlack, True
            End If
        ElseIf Col = 2 + (4 * A) Then
            If NulosC(fg1.TextMatrix(Row, Col)) = "" Then
                fg1.TextMatrix(Row, Col - 1) = 0
                fg1.TextMatrix(Row, Col + 1) = ""
                fg1.TextMatrix(Row, Col + 2) = 0
            
                FORMATO_CELDA fg1, Row, Col - 1, vbBlack, False
                FORMATO_CELDA fg1, Row, Col - 2, vbBlack, False
                FORMATO_CELDA fg1, Row, Col - 3, vbBlack, False
                
            End If
        
        End If
        
        
    Next

End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        fg1.Editable = flexEDNone
    Else
        fg1.Editable = flexEDKbdMouse
    End If
    
End Sub

Private Sub Form_Activate()
    'Modificado: 11/01/11 Johan Castro
    '            Agregar linea de codigo para bloquear accesos de usuarios


    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        
        pCargarGrid
        
    End If

End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
End Sub


Sub Blanquea()
    
    TxtDescripcion.Text = ""
    TxtColumna.Text = ""
    
    fg1.Rows = 1
End Sub

Sub Bloquea(band As Boolean)

    TxtDescripcion.Locked = Not band
    TxtColumna.Locked = Not band
    
    

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
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Informe ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset '--relacionado a los aportes que se consideraran
    Dim RstTmp As New ADODB.Recordset '--temporal
    Dim nSQl As String
    Dim xId As Double
    Dim A&, B&, mCorr&
    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("con_informe", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_informe", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_informe WHERE id = " & xId & ";", xCon
        
        '--eliminar el detalle del informe
        xCon.Execute "DELETE FROM con_informedet WHERE idinf = " & xId & ";"
        
    End If
    
    '******************
    mIdRegistro = xId
    '******************
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_informedet", xCon
    
    
    
    RstCab("descripcion") = NulosC(TxtDescripcion.Text)
    RstCab("cancol") = NulosN(TxtColumna.Text)
    
    RstCab.Update
    
    
    '--
    '--grabar el detalle del informe
    
    mCorr = 1
    For A = 0 To NulosN(TxtColumna.Text) - 1
        
        For B = 1 To fg1.Rows - 1
            RstDet.AddNew
            RstDet("idinf") = xId
            RstDet("corr") = mCorr
            RstDet("posicion") = A + 1
            RstDet("idcpto") = NulosN(fg1.TextMatrix(B, 1 + (4 * A)))
            RstDet("descripcion") = NulosC(fg1.TextMatrix(B, 2 + (4 * A)))
            RstDet("negrita") = NulosN(fg1.TextMatrix(B, 4 + (4 * A)))
            
            RstDet("ancho") = fg1.ColWidth(2 + (4 * A))
            
            RstDet.Update
            mCorr = mCorr + 1
        Next B

    Next A
    xCon.CommitTrans
    Set RstCab = Nothing
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    MsgBox "Los datos del Informe " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo


    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar al Personal por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub Nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    
    
    Label5.Caption = "Agregando Informe"

    
    xHorIni = Time

    '-------------------------------------------
    TxtDescripcion.SetFocus
    
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Informe"

    ActivaTool
    
    QueHace = 2
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
    End If
    
    Bloquea True
    
    TabOne1.TabEnabled(0) = False
    
    Agregando = False
    
    xHorIni = Time
    
    TxtDescripcion.SetFocus

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
    Dim nSQl As String
    
    Dim RstBus As New ADODB.Recordset
    xId = RstFrm.Fields("id")
    
    '---------------------------------------------------------------------------------
    
    Rpta = MsgBox("Esta seguro de eliminar el Informe seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
    
        xCon.Execute "DELETE FROM con_informedet WHERE idinf = " & xId & ";" '--replaciona al cuentas contables
        xCon.Execute "DELETE FROM con_informe WHERE id = " & xId & ";"
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo


        MsgBox "El Informe se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then Cancelar
    If Button.Index = 6 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg1.Refresh
            If RstFrm.State = 1 Then
                If RstFrm.RecordCount <> 0 Then
                    RstFrm.MoveFirst
                    RstFrm.Find "id =" & mIdRegistro
                    If RstFrm.EOF = True Then RstFrm.MoveFirst
                End If
            End If
            Cancelar
        End If
    End If

    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstFrm.Filter = adFilterNone
    End If
''    If Button.Index = 10 Then Buscar
    
''    If Button.Index = 12 Then pExportar
''    If Button.Index = 13 Then pImorimir

    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Buscar()
    
End Sub

Sub Filtrar()

End Sub


'****************************************************************************************

Sub MuestraSegundoTab()

    Dim QueHaceTmp As Integer
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    
    
    TxtDescripcion.Text = RstFrm("descripcion")
    
    TxtColumna.Text = NulosN(RstFrm("cancol"))
    If NulosN(TxtColumna.Text) = 0 Then TxtColumna.Text = 1
    '--dar  formato al grid
    FormatoGrid
    '----------------------------------
    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    Dim A, B, C As Long
    '--obtener la consulta del detalle del informe
    nSQl = "SELECT con_informedet.idcpto, con_informedet.corr, con_informedet.posicion, con_informedet.idcpto, con_informedet.descripcion, con_informedet.negrita, con_concepto.formula, con_concepto.variable,con_informedet.ancho " _
        + vbCr + " FROM con_informedet LEFT JOIN con_concepto ON con_informedet.idcpto = con_concepto.id " _
        + vbCr + " WHERE (((con_informedet.idinf) = " & NulosN(RstFrm("id")) & ")) " _
        + vbCr + " ORDER BY con_informedet.corr;"
    
    RST_Busq RstTmp, nSQl, xCon
    
    DoEvents
    
    '--recorrer la cantidad de filas del grid
    For A = 0 To NulosN(TxtColumna.Text) - 1
        '--hacer el filtro para agregar el contenido del rst al grid
        RstTmp.Filter = "posicion=" & A + 1
        If RstTmp.RecordCount <> 0 Then
            If fg1.Rows - 1 < RstTmp.RecordCount Then
                fg1.Rows = RstTmp.RecordCount + 1
            End If
            RstTmp.MoveFirst
        End If
        '--posicionar en la primera fila
        C = 1
        '--agregando al grid
        Do While Not RstTmp.EOF
            fg1.TextMatrix(C, 1 + (4 * A)) = RstTmp("idcpto")
            fg1.TextMatrix(C, 2 + (4 * A)) = NulosC(RstTmp("descripcion"))
            fg1.TextMatrix(C, 3 + (4 * A)) = NulosC(RstTmp("variable"))
            fg1.TextMatrix(C, 4 + (4 * A)) = NulosN(RstTmp("negrita"))
            
            '--dar ancho a la columna
            fg1.ColWidth(2 + (4 * A)) = NulosN(RstTmp("ancho"))
            
            '--incrementar la fila
            C = C + 1
         RstTmp.MoveNext
        Loop
    Next A
    
    Set RstTmp = Nothing
    '----------------------------------

End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQl As String
    
    Set RstFrm = Nothing
    
    TDB_FiltroLimpiar Dg1
    
    RstFrm.Filter = adFilterNone
    DoEvents

    nSQl = "SELECT con_informe.* FROM con_informe;"

    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQl, xCon
    
    Set Dg1.DataSource = RstFrm
    TabOne1.CurrTab = 0
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Function fValidarDatos() As Boolean
    '===================================================================================================
    'Creado : 29/09/08 Por: Johan Castro
    'Propósito: validar antes de grabar el registro
    '
    'Entradas:  Ninguno
    '
    'Resultados: true:  Listo para grabar
    '            false: Mensaje de alerta, tiene que corregir para continuar
    '
    'Otros:
    '
    'Modificado :
    
    '===================================================================================================
    
    If NulosC(TxtDescripcion.Text) = "" Then
        MsgBox "Falta especificar la Descripción del Informe", vbExclamation, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If
    
    '--verificar que la variable no tenga cietos caracteres
    If fg1.Rows = fg1.FixedRows Then
        MsgBox "Falta agregar el detalle del Informe", vbExclamation, xTitulo
        Exit Function
    End If
    
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    If QueHace = 1 Then
        nSQl = "SELECT con_informe.descripcion FROM con_informe WHERE (((UCase([con_informe].[descripcion]))='" & UCase(TxtDescripcion.Text) & "'));"
    Else
        nSQl = "SELECT con_informe.descripcion fROM con_informe WHERE (((UCase([con_informe].[descripcion]))='" & UCase(TxtDescripcion.Text) & "') AND ((con_informe.id)<>" & NulosN(RstFrm.Fields("id")) & "));"
    End If
    RST_Busq RstTmp, nSQl, xCon
    If RstTmp.RecordCount <> 0 Then
        MsgBox "Existe un Informe que tiene asignado la misma descripción", vbExclamation, xTitulo
        Exit Function
    End If
    '--
    fValidarDatos = True
    
End Function


'************************************************************
'************************************************************

Private Sub cmd_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0 '--ADD REG
            If fg1.Row <= fg1.FixedRows Or fg1.Row = fg1.Rows - 1 Then
                fg1.AddItem ""
            Else
                fg1.AddItem "", fg1.Row
            End If
        Case 1 '--DEL REG
            If fg1.Row <= 0 Then
                MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una fila correcta", vbExclamation, xTitulo
                Exit Sub
            End If
    
            fg1.RemoveItem (fg1.Row)
            
    End Select
End Sub

Private Sub TxtColumna_KeyPress(KeyAscii As Integer)
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtColumna_Validate(Cancel As Boolean)
    FormatoGrid
End Sub

Private Sub FormatoGrid()
    '===================================================================================================
    'Creado : 29/09/08 Por: Johan Castro
    'Propósito: procedimiento para dar formato al grid cuando cambie la cantidad de columnas
    '
    'Entradas:  Ninguno
    '
    'Resultados: Grid listo para agregar registros y asociar los conceptos
    '
    'Otros: Es el inicio del proceso de generacion del informe
    '       Depende del numero de columnas que ingresa el usuario,si no ingresa su valor sera 1 por defecto
    '
    'Modificado :
    
    '===================================================================================================
    
    
    Dim A As Integer
        
    '--evaluar si la cantdad de filas es inferior a 1
    If NulosN(TxtColumna.Text) < 1 Then
        MsgBox "Cantidad de Columnas como mínimo debe tomar valor = 1", vbInformation, xTitulo
        TxtColumna.Text = 1
        Exit Sub
    End If
    
    '--inciarlizar las filas y columnas par luego dar formato
    fg1.Rows = 1
    fg1.Cols = 1
    '--dar cantida de columnas, esta depende de la cantidad de columnas
    
    fg1.Cols = 4 * NulosN(TxtColumna.Text) + 1
    For A = 0 To NulosN(TxtColumna.Text) - 1
        fg1.TextMatrix(0, 1 + (4 * A)) = "Id"
        fg1.TextMatrix(0, 2 + (4 * A)) = "Descripción"
        fg1.TextMatrix(0, 3 + (4 * A)) = "Variable"
        fg1.TextMatrix(0, 4 + (4 * A)) = "Negrita"
        '--estableciendo el formato
        GRID_COMBOLIST fg1, 2 + (4 * A)
        
        fg1.ColAlignment(3 + (4 * A)) = flexAlignLeftCenter
        fg1.ColAlignment(4 + (4 * A)) = flexAlignCenterCenter
        fg1.ColDataType(4 + (4 * A)) = flexDTBoolean
        
        fg1.ColWidth(1 + (4 * A)) = 0
        fg1.ColWidth(2 + (4 * A)) = 3480
        fg1.ColWidth(3 + (4 * A)) = 960
        fg1.ColWidth(4 + (4 * A)) = 600
        
    Next



End Sub
