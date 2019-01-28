VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManUsuarioAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Grupos"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7065
      Top             =   -15
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
            Picture         =   "FrmManUsuarioAcceso.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarioAcceso.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6870
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   8115
      _cx             =   14314
      _cy             =   12118
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
         Height          =   6450
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   8025
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5895
            Left            =   30
            TabIndex        =   6
            Top             =   450
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   10398
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "idusu"
            Columns(0).DataField=   "idusu"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Usuario"
            Columns(1).DataField=   "desusu"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Número de Serie"
            Columns(2).DataField=   "numser"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=8017"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7938"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2831"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2752"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(52)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(55)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   "Named:id=36:Selected"
            _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=37:Caption"
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Acceso"
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
            Left            =   15
            TabIndex        =   2
            Top             =   30
            Width           =   7965
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6450
         Left            =   8760
         TabIndex        =   3
         Top             =   375
         Width           =   8025
         Begin VB.Frame Frame3 
            Caption         =   "[ Detalles de Acceso ]"
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
            Height          =   4755
            Left            =   30
            TabIndex        =   15
            Top             =   1650
            Width           =   7935
            Begin VB.Frame Frame9 
               Height          =   4380
               Left            =   6450
               TabIndex        =   17
               Top             =   270
               Width           =   1425
               Begin VB.CommandButton cmdPer 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  TabIndex        =   21
                  TabStop         =   0   'False
                  ToolTipText     =   "Eliminar Personal"
                  Top             =   1140
                  Width           =   1290
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Agregar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  TabIndex        =   20
                  ToolTipText     =   "Agregar Personal"
                  Top             =   165
                  Width           =   1290
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Seleccionar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  TabIndex        =   19
                  ToolTipText     =   "Agregar Personal"
                  Top             =   510
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.CommandButton cmdPer 
                  Caption         =   "Eliminar &Todos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  TabIndex        =   18
                  TabStop         =   0   'False
                  ToolTipText     =   "Eliminar Personal"
                  Top             =   1500
                  Width           =   1290
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   4305
               Left            =   60
               TabIndex        =   22
               Top             =   360
               Width           =   6315
               _cx             =   11139
               _cy             =   7594
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
               Rows            =   2
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManUsuarioAcceso.frx":2B10
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
            Begin VB.Label Label4 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "LblIdRec"
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
               Left            =   9600
               TabIndex        =   16
               Top             =   330
               Width           =   1185
            End
         End
         Begin VB.Frame FrmReceta 
            Caption         =   "[ Detalles de Usuario ]"
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
            Height          =   1185
            Left            =   30
            TabIndex        =   7
            Top             =   390
            Width           =   7935
            Begin VB.TextBox txtNumSer 
               Height          =   285
               Left            =   1140
               TabIndex        =   14
               Text            =   "txtNumSer"
               Top             =   690
               Width           =   6705
            End
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2130
               Picture         =   "FrmManUsuarioAcceso.frx":2B6E
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   350
               Width           =   225
            End
            Begin VB.TextBox TxtIdUsu 
               Height          =   300
               Left            =   1140
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   9
               Text            =   "TxtIdUsu"
               Top             =   315
               Width           =   1245
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Núm. Serie"
               Height          =   195
               Index           =   1
               Left            =   210
               TabIndex        =   13
               Top             =   750
               Width           =   780
            End
            Begin VB.Label lblUsuario 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblUsuario"
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
               Left            =   2400
               TabIndex        =   12
               Top             =   315
               Width           =   5445
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Usuario"
               Height          =   195
               Index           =   0
               Left            =   210
               TabIndex        =   11
               Top             =   360
               Width           =   540
            End
            Begin VB.Label LblIdRec 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "LblIdRec"
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
               Left            =   9600
               TabIndex        =   10
               Top             =   330
               Width           =   1185
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Acceso"
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
            Height          =   210
            Left            =   30
            TabIndex        =   4
            Top             =   60
            Width           =   7905
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmManUsuarioAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANGRUPOS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA LA CREACION Y MANTENIMIENTO DE GRUPOS DE TRABAJADORES
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 05/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer             ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean           ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim RstFrm As New ADODB.Recordset  ' RECORDSET QUE ALAMCENARA LOS DATOS DE LA TABLA pro_grupo
Dim Agregando As Boolean           ' VARIABLE QUE INDICA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID
Dim fOrdenLista As Boolean         ' especfica el orden de la lista de la consulta
Dim fSeleccionVarios As Boolean
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDENTE LAS COLUMNAS DEL CONTROL Dg1
On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 1 Or Fg1.Col = 3 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 3 Then Exit Sub
    
    If Col <> 3 Then Exit Sub
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 114 Or KeyCode = 45 Then cmdPer_click 0    'F3 = Agrega Item
    
    If KeyCode = 115 Or KeyCode = 46 Then cmdPer_click 1    'F4 = Eliminar Item
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then PopupMenu Menu1
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim xCodTar As Double
    Dim nSQL As String
    Dim nSQLId As String
    On Error GoTo error
    
    If Col = 1 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        Dim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "nombres":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Area":                 xCampos(1, 1) = "area":         xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
        
        ' si hay registros anteriomente seleccionadas no considerar de nuevo
        nSQLId = GRID_GENERAR_SQL_ID(Fg1, 4, " and pro_emp.id", "NOT IN")
        nSQL = "SELECT pla_empleados.id AS idemp, 0 AS numgrupo, [pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_area.descripcion AS area, pro_emp.id AS idper " _
            + vbCr + " FROM ((mae_area RIGHT JOIN pla_empleados ON mae_area.id = pla_empleados.idarea) INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
            + vbCr + " WHERE pla_empleados.fchcese is null and (((pro_emp.id ) Not In (select pro_grupodet.idper from pro_grupodet))) and pro_empdet.idfun = 6 " & nSQLId _
            + vbCr + " ORDER BY [pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
        
        If fSeleccionVarios = True Then
            CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Personal"
        Else
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Personal", "nombres", "nombres", Principio
        End If
        
        Agregando = True
        
        If xRs.State = 1 Then
            If fSeleccionVarios = True Then xRs.MoveFirst
            Do While Not xRs.EOF
                xCodTar = NulosN(Fg1.TextMatrix(Fg1.Row, 4))
                Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("nombres"))
                Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("area"))
                Fg1.TextMatrix(Fg1.Row, 3) = 0
                Fg1.TextMatrix(Fg1.Row, 4) = NulosN(xRs("idper"))
                If fSeleccionVarios = False Then
                    Exit Do
                Else
                    xRs.MoveNext
                    If xRs.EOF = False Or xRs.BOF = False Then
                        Fg1.Rows = Fg1.Rows + 1
                        Fg1.Row = Fg1.Rows - 1
                    End If
                End If
            Loop
        End If
        Set xRs = Nothing
    End If
    
    Agregando = False
    Fg1.Col = 3
    Fg1.SetFocus
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick (" & Row & "," & Col & ")"
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUADO SE CARGUE EL FORMULARIO
    If SeEjecuto = True Then Exit Sub
    Dim Rpta As Integer
    SeEjecuto = False
    Fg1.ColWidth(4) = 0
    pCargarGrid
    SeEjecuto = True
    
    '--Almacenar temporalmente el codigo del menu
    IdMenuActivo = xIdMenu
    '--bloquear accesos
    OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
    'ocultar siempre el boton nuevo
    Toolbar1.Buttons(1).Visible = False
    
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA TABLA pro_grupo EN EL CONTROL Dg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    Dim nSQL  As String
    
    nSQL = "SELECT pla_empleados.id AS idemp, pro_grupo.num AS numgrupo, UCase([pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) AS nombres, mae_area.descripcion AS area, pro_emp.id AS idper " _
        + vbCr + " FROM pro_grupo INNER JOIN (((mae_area RIGHT JOIN pla_empleados ON mae_area.id = pla_empleados.idarea) INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_grupodet ON pro_emp.id = pro_grupodet.idper) ON pro_grupo.id = pro_grupodet.idgrupo " _
        + vbCr + " ORDER BY pro_grupo.num, UCase([pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]);"

    ' cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    CentrarFrm Me
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    Dg1.BatchUpdates = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Habilitar_Obj False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg1.DataSource = Nothing
End Sub

Private Sub Menu1_1_Click()
    cmdPer_click 0
End Sub

Private Sub Menu1_3_Click()
    cmdPer_click 1
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then RstFrm.Filter = ""
    
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 13 Then pExportar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_grupo
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        If RstFrm.RecordCount = 0 Then
            MsgBox "No hay registros", vbExclamation, xTitulo
        Else
            MsgBox "Seleccione un Registro para Eliminar", vbExclamation, xTitulo
        End If
        Exit Sub
    End If
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    TabOne1.CurrTab = 0
    
    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETe * FROM pro_grupodet WHERE idgrupo = " & RstFrm("numgrupo") & " and idper = " & RstFrm("idper")
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        Dg1.Refresh
        
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No ha Configurado ningún Grupo, ¿Desea Configurarlo ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                Modificar
            Else
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle del Grupo"
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
   
    QueHace = 2
    ActivaTool
    Habilitar_Obj True
    MuestraSegundoTab
    GRID_COMBOLIST Fg1, 1
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    Label1.Caption = "Modificando Grupo"
    cmdPer(0).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    On Error GoTo error
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    ' limpiar controles
    Blanquea
    
    nSQL = "SELECT pla_empleados.id AS idemp, pro_grupo.num AS numgrupo, UCase([pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) AS nombres, mae_area.descripcion AS area, pro_emp.id AS idper " _
            + vbCr + " FROM pro_grupo INNER JOIN (((mae_area RIGHT JOIN pla_empleados ON mae_area.id = pla_empleados.idarea) INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_grupodet ON pro_emp.id = pro_grupodet.idper) ON pro_grupo.id = pro_grupodet.idgrupo " _
            + vbCr + " ORDER BY pro_grupo.num, UCase([pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]);"

    RST_Busq RstTmp, nSQL, xCon

    With RstTmp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(.Fields("nombres"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(.Fields("area"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(.Fields("numgrupo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(.Fields("idper"))
            .MoveNext
        Loop
        GRID_AGRUPAR Fg1, 3
    End With
    
    Set RstTmp = Nothing
    Exit Sub

error:
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

'*****************************************************************************************************
'* Nombre           : Habilitar_Obj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O ESACTIVA LOS CONTROLES TEXTBOX Y COMMAND DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Habilitar_Obj(band As Boolean)
    habilitar cmdPer, band
    TabOne1.CurrTab = IIf(band = False, 0, 1)
    TabOne1.TabEnabled(0) = Not band
    
    If band = False Then
        Fg1.SelectionMode = flexSelectionByRow
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA LAS FILAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Blanquea()
    Fg1.Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Nuevo()
    QueHace = 1
    ActivaTool
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregando Grupo"
    GRID_COMBOLIST Fg1, 1
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_grupo, ESTA FUNCION DEVUELEVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICA QUE LOS DATOS INGRESADOS SON LOS CORRECTOS
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xCod&, xCol&, xFil&
    Dim nSQL As String
    
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    xCon.BeginTrans
    
    RST_Busq RstCab, "SELECT top 1 * FROM pro_grupo ", xCon
    RST_Busq RstDet, "SELECT top 1 * FROM pro_grupodet", xCon
    
    xCon.Execute "DELETE * FROM pro_grupodet "
    xCon.Execute "DELETE * FROM pro_grupo "
    
    For xFil = 1 To Fg1.Rows - 1
        DoEvents
        Set RstTmp = Nothing
        RST_Busq RstTmp, "select pro_grupo.id from pro_grupo where pro_grupo.num =  " & NulosN(Fg1.TextMatrix(xFil, 3)), xCon
        
        If RstTmp.RecordCount = 0 Then
            xCod = HallaCodigoTabla("pro_grupo", xCon, "id")
            RstCab.AddNew
            RstCab("id") = xCod
            RstCab("num") = NulosN(Fg1.TextMatrix(xFil, 3))
        Else
            xCod = RstTmp("id")
        End If
        
        RstCab.Update
        
        RstDet.AddNew
        RstDet("idgrupo") = xCod
        RstDet("idper") = NulosN(Fg1.TextMatrix(xFil, 4))
        RstDet.Update
    Next
    
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    xCon.CommitTrans
    Grabar = True

SALIR:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstTmp = Nothing
    Me.MousePointer = vbDefault
    Exit Function

LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA QUE EL CONTROL Fg1 TENGA DATOS, ESTA FUNCION DEVUELVE VERDADERO SI LOS
'*                    DATOS EN EL CONTROL Fg1 SON CORRECTOS
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    Dim mRow&, QGrid&
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay registros para grabar" + vbCr + "Es necesario que  contenga por lo menos un registro", vbExclamation, xTitulo
        Exit Function
    End If
    
    fValidarDatos = True
End Function
 
'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET xRs
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "nombres":      xCampos(0, 2) = "5000":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Grupo":             xCampos(1, 1) = "numgrupo":     xCampos(1, 2) = "1000":     xCampos(1, 3) = "N"
            
    nSQL = "SELECT pla_empleados.id AS idemp, pro_grupo.num AS numgrupo, UCase([pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) AS nombres, mae_area.descripcion AS area, pro_emp.id AS idper " _
            + vbCr + " FROM pro_grupo INNER JOIN (((mae_area RIGHT JOIN pla_empleados ON mae_area.id = pla_empleados.idarea) INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_grupodet ON pro_emp.id = pro_grupodet.idper) ON pro_grupo.id = pro_grupodet.idgrupo " _
            + vbCr + " ORDER BY pro_grupo.num, UCase([pla_empleados.apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]);"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Personal", "nombres", "nombres", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))

SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO SOBRE EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Filtrar()
    ReDim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Apellidos y Nombres": xCampos(0, 1) = "nombres":    xCampos(0, 2) = "C":     xCampos(0, 3) = "5000"
    xCampos(1, 0) = "Nº Grupo":            xCampos(1, 1) = "numgrupo":   xCampos(1, 2) = "N":     xCampos(1, 3) = "900"
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
End Sub

'*****************************************************************************************************
'* Nombre           : pBuscarVSFlexGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA UN DATO EN EL CONTRO FLEXGRID
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pBuscarVSFlexGrid()
    On Error GoTo error
    
    Dim xExport As New SGI2_funciones.formularios
    Dim xCampos(0, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Apellidos y Nombres":     xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    
    xExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos()
    Set xExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub

Private Sub cmdPer_click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            fSeleccionVarios = False
            pRegistroAdd
        
        Case 1 '--eliminar
            pRegistroDel
        
        Case 2 '--ordenar
            GRID_ORDENAR Fg1, 1, 3
            GRID_AGRUPAR Fg1, 3
        
        Case 3 '--buscar
            pBuscarVSFlexGrid
        
        Case 4 '--seleccionar
            fSeleccionVarios = True
            pRegistroAdd
    End Select
End Sub

Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > Fg1.FixedRows Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 4)) = 0 Then
            MsgBox "Seleccione un Personal", vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fInsertar = True Then Fg1.AddItem ""
    
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 1
    
    If fInsertar = True Then Fg1_CellButtonClick Fg1.Rows - 1, 1
    
    Fg1.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
    If Fg1.Rows > 1 Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 1
        Fg1.SetFocus
    Else
        cmdPer(0).SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL RECORDSET RSTTMP
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0
    
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
        
    Dim xCampos(2, 3) As String
    
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Nº Grupo":              xCampos(0, 1) = "numgrupo":   xCampos(0, 2) = 2:  xCampos(0, 3) = "900"
    xCampos(1, 0) = "Apellidos y Nombres":   xCampos(1, 1) = "nombres":    xCampos(1, 2) = 0:  xCampos(1, 3) = "4500"
    xCampos(2, 0) = "Area":                  xCampos(2, 1) = "area":       xCampos(2, 2) = 0:  xCampos(2, 3) = "1500"
    Set RstTmp = RstFrm.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Listado de Grupos de Producción", "", "", "Relación de Grupos", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub


