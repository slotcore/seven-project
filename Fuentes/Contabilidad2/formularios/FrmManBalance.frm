VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManBalance 
   Caption         =   "Contabilidad - Configuración Estados Financieros - Balance"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11685
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManBalance.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
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
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
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
      Height          =   6675
      Left            =   15
      TabIndex        =   0
      Top             =   375
      Width           =   11670
      _cx             =   20585
      _cy             =   11774
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
         Height          =   6255
         Left            =   45
         TabIndex        =   3
         Top             =   375
         Width           =   11580
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5865
            Left            =   30
            TabIndex        =   15
            Top             =   360
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   10345
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
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Observación"
            Columns(2).DataField=   "observacion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=6562"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6482"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=11324"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=11245"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=34,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=37:Caption"
            _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(59)  =   "Named:id=38:HighlightRow"
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Balance"
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
            TabIndex        =   4
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   12315
         TabIndex        =   5
         Top             =   375
         Width           =   11580
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   4560
            Left            =   0
            TabIndex        =   11
            Top             =   750
            Width           =   11565
            _cx             =   20399
            _cy             =   8043
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   8388608
            Caption         =   "        [Activo]        |        [Pasivo]        "
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
            Begin VB.Frame fr 
               BorderStyle     =   0  'None
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
               Height          =   4185
               Index           =   0
               Left            =   12210
               TabIndex        =   13
               Top             =   330
               Width           =   11475
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   3285
                  Index           =   1
                  Left            =   15
                  TabIndex        =   18
                  Top             =   195
                  Width           =   7275
                  _cx             =   12832
                  _cy             =   5794
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManBalance.frx":277E
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
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   3285
                  Index           =   3
                  Left            =   7350
                  TabIndex        =   19
                  Top             =   195
                  Width           =   4095
                  _cx             =   7223
                  _cy             =   5794
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
                  FormatString    =   $"FrmManBalance.frx":286A
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
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   2
                  Left            =   15
                  TabIndex        =   28
                  Top             =   3405
                  Width           =   7260
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   3
                     Left            =   1305
                     TabIndex        =   30
                     ToolTipText     =   "Eliminar Registro"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   2
                     Left            =   60
                     TabIndex        =   29
                     ToolTipText     =   "Agregar Registro"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   3
                  Left            =   7350
                  TabIndex        =   23
                  Top             =   3405
                  Width           =   4080
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   9
                     Left            =   2835
                     TabIndex        =   33
                     ToolTipText     =   "Eliminar Cuenta"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   7
                     Left            =   60
                     TabIndex        =   32
                     ToolTipText     =   "Agregar Cuenta Contable"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Seleccionar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   8
                     Left            =   1305
                     TabIndex        =   31
                     ToolTipText     =   "Agregar Cuentas Contables"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VB.Label lbl_cabecera 
                  Caption         =   "lbl_cabecera(1)"
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
                  Height          =   255
                  Index           =   1
                  Left            =   90
                  TabIndex        =   34
                  Top             =   3960
                  Width           =   11340
               End
            End
            Begin VB.Frame fr 
               BorderStyle     =   0  'None
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
               Height          =   4185
               Index           =   1
               Left            =   45
               TabIndex        =   12
               Top             =   330
               Width           =   11475
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   3285
                  Index           =   0
                  Left            =   15
                  TabIndex        =   16
                  Top             =   195
                  Width           =   7275
                  _cx             =   12832
                  _cy             =   5794
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManBalance.frx":28FE
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
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   3285
                  Index           =   2
                  Left            =   7350
                  TabIndex        =   17
                  Top             =   195
                  Width           =   4095
                  _cx             =   7223
                  _cy             =   5794
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
                  FormatString    =   $"FrmManBalance.frx":29F2
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
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   0
                  Left            =   15
                  TabIndex        =   20
                  Top             =   3405
                  Width           =   7260
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   1
                     Left            =   1305
                     TabIndex        =   22
                     ToolTipText     =   "Eliminar Registro"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   0
                     Left            =   60
                     TabIndex        =   21
                     ToolTipText     =   "Agregar Registro"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   1
                  Left            =   7350
                  TabIndex        =   24
                  Top             =   3405
                  Width           =   4080
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   6
                     Left            =   2835
                     TabIndex        =   27
                     ToolTipText     =   "Eliminar Cuenta"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   4
                     Left            =   60
                     TabIndex        =   26
                     ToolTipText     =   "Agregar Cuenta Contable"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Seleccionar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   5
                     Left            =   1305
                     TabIndex        =   25
                     ToolTipText     =   "Agregar Cuentas Contables"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VB.Label lbl_cabecera 
                  Caption         =   "lbl_cabecera(0)"
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
                  Height          =   240
                  Index           =   0
                  Left            =   90
                  TabIndex        =   35
                  Top             =   3960
                  Width           =   11355
               End
            End
         End
         Begin VB.TextBox txt 
            Height          =   555
            Index           =   2
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Tag             =   "null"
            Text            =   "FrmManBalance.frx":2A86
            Top             =   5565
            Width           =   11505
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1215
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txt(1)"
            Top             =   360
            Width           =   5895
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   0
            Left            =   10200
            TabIndex        =   8
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   345
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Observación"
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   10
            Top             =   5355
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   9600
            TabIndex        =   7
            Top             =   465
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Balance"
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
            Left            =   15
            TabIndex        =   6
            Top             =   75
            Width           =   11550
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
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu Menu3_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu3_2 
         Caption         =   "Seleccionar"
      End
      Begin VB.Menu Menu3_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu3_4 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Menu3_5 
         Caption         =   "Eliminar Todo"
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "Menu4"
      Visible         =   0   'False
      Begin VB.Menu Menu4_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu4_2 
         Caption         =   "Seleccionar"
      End
      Begin VB.Menu Menu4_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu4_4 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Menu4_5 
         Caption         =   "Elimianar Todos"
      End
   End
End
Attribute VB_Name = "FrmManBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANBALANCE.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE LAS ALTAS Y BAJAS EN LA TABLA con_balance
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 27/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer                  ' VARIABLE QUE INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean                ' VARIABLE PARA CONTROLAR QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim RstFrm As New ADODB.Recordset       ' RECORDSET QUE ALACENARA LOS DATOS DE LA TABLA con_balance
Dim Agregando As Boolean                ' VARIABLE QUE INDICA QUE SE ESTA AGREGANDO UN FILA A LOS CONTROLS FLEXGRID
Dim M_MES_ACTIVO  As Integer            ' INDICA EL MES ACTIVO
Dim TmpCta As New ADODB.Recordset
Dim IdFila As Integer                   ' INDICA LA FILA DE LA CABECERA DEL BALANCE
                                        ' UTIL PARA DIFERENCIAR EL DETALLE DEL BALANCE (TIPO:[ACTIVO,PASIVO]; IdFila)
Dim IDBALANCE_CABECERA As Long          ' INDICA EL ULTIMO ID DE LA CABECERA DEL BALANCE

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UNA FILA AL CONTROL Fg
'* Paranetros       : NOMBRE      |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Index       |  Integer    |  ESPECIFICA EL INDICE DEL CONTROL Fg
'*                    fSelVarios  |  Boolean    |  ESPECIFICA SI SE MUESTRA TODO O SOLO EL SELECCIONADO
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd(Index As Integer, Optional fSelVarios As Boolean = False)
    If QueHace = 3 Then Exit Sub
    
    If txt(1).Text = "" Then
        MsgBox "Ingrese la Descripción del Balance", vbExclamation, xTitulo
        txt(1).SetFocus
        Exit Sub
    End If
    
    Agregando = True
    Select Case Index
        Case 0, 1 '--
            If fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "" Or fg(Index).TextMatrix(fg(Index).Rows - 1, 4) = "" Then
                MsgBox "Falta ingresar " + IIf(fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "", "el N° de Orden", "la descripción") _
                + vbCr + "Ingrese el dato requerido para agregar un nuevo registro", vbExclamation, xTitulo
                
                fg(Index).Col = IIf(fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "", 3, 4)
                GoTo SALIR
            End If
        
            fg(Index).AddItem ""
            
            fg(Index).TextMatrix(fg(Index).Rows - 1, 1) = IdFila
            fg(Index).TextMatrix(fg(Index).Rows - 1, 2) = IDBALANCE_CABECERA
            
            IdFila = IdFila + 1:        IDBALANCE_CABECERA = IDBALANCE_CABECERA + 1
            fg(Index).Row = fg(Index).Rows - 1
            fg(Index).Col = IIf(fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "", 3, 4)
            
            GoTo SALIR
        
        Case 2, 3
            If fg(Index - 2).TextMatrix(fg(Index - 2).Row, 3) = "" Or fg(Index - 2).TextMatrix(fg(Index - 2).Row, 4) = "" Then
                MsgBox "Falta ingresar datos en la Cabecera" + _
                vbCr + "Información Requerida: " + IIf(fg(Index - 2).TextMatrix(fg(Index - 2).Row, 3) = "", "N° de Orden", "Descripción"), vbExclamation, xTitulo
                
                fg(Index - 2).Col = IIf(fg(Index - 2).TextMatrix(fg(Index - 2).Rows - 1, 3) = "", 3, 4)
                GoTo SALIR:
            End If
    End Select
    
    ' GENERAR EL WHERE DE LOS ID'S DE CUENTA PARA QUE NO SE REPITAN
    Dim SQL_ID As String
    Dim RstCtaTmp As New ADODB.Recordset
    
    RST_Busq RstCtaTmp, "select * from con_planctas order by cuenta asc", xCon
    
    SQL_ID = ""
    If RstCtaTmp.State = 0 Then GoTo SALIR
    If RstCtaTmp.EOF = True Or RstCtaTmp.BOF = True Or RstCtaTmp.RecordCount = 0 Then GoTo SALIR
    TmpCta.Filter = ""
    If TmpCta.EOF = False Or TmpCta.BOF = False Or TmpCta.RecordCount <> 0 Then TmpCta.MoveFirst
    
    Do While Not TmpCta.EOF
        RstCtaTmp.Filter = "cuenta='" + TmpCta.Fields("cuenta") + "'"
        
        If RstCtaTmp.RecordCount <> 0 Then
            If NulosN(RstCtaTmp.Fields("dissegsal")) = "-1" Then
                If (Index = 2 And TmpCta.Fields("tipo") = 1) Or (Index = 3 And TmpCta.Fields("tipo") = 2) Then  '--ACTIVO
                    SQL_ID = SQL_ID + CStr(TmpCta.Fields("idcuenta")) + ","
                End If
            Else
                SQL_ID = SQL_ID + CStr(TmpCta.Fields("idcuenta")) + ","
            End If
        Else
            SQL_ID = SQL_ID + CStr(TmpCta.Fields("idcuenta")) + ","
        End If
        TmpCta.MoveNext
    Loop
    
    Set RstCtaTmp = Nothing
    
    If SQL_ID <> "" Then SQL_ID = " WHERE con_planctas.id NOT IN (" + Left(SQL_ID, Len(SQL_ID) - 1) + ") "

    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    Dim xCampos(2, 5) As String
    Dim nSQL As String
    
    xCampos(0, 0) = "N° Cta.":      xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "6500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    
    ' TIPO,BAL_CAB,IDCUENTA,NUMCUENTA,DESC_CUENTA
    If fSelVarios = True Then
        ' Mostrar las cuentas que tienen movimiento
        nSQL = "SELECT " + CStr(Index - 1) + " as tipo, " + fg(Index - 2).TextMatrix(fg(Index - 2).Row, 1) + " as orden," + fg(Index - 2).TextMatrix(fg(Index - 2).Row, 2) + " as idbal, con_planctas.id as idcuenta, con_planctas.cuenta & ' ' AS cuenta, con_planctas.descripcion " _
            + vbCr + " FROM con_diario INNER JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
            + vbCr + SQL_ID _
            + vbCr + " GROUP BY   con_planctas.id, con_planctas.cuenta & ' ', con_planctas.descripcion, con_planctas.cuenta " _
            + vbCr + " ORDER BY con_planctas.cuenta ASC "

        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Cuentas con Movimientos"
    Else
        ' Mostrar todas las cuentas
        nSQL = "SELECT " + CStr(Index - 1) + " as tipo, " + fg(Index - 2).TextMatrix(fg(Index - 2).Row, 1) + " as orden," + fg(Index - 2).TextMatrix(fg(Index - 2).Row, 2) + " as idbal, con_planctas.id as idcuenta, con_planctas.cuenta & ' ' AS cuenta, con_planctas.descripcion " _
            + vbCr + " FROM con_planctas " _
            + vbCr + SQL_ID _
            + vbCr + " ORDER BY con_planctas.cuenta ASC "
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Agregando Cuentas", "cuenta", "cuenta", Principio
    End If
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    ' CARGANDO A UN TEMPORAL
    If fSelVarios = True Then
        xRs.MoveFirst
        CARGAR_RST_TMP TmpCta, xRs
    Else
        CARGAR_RST_TMP TmpCta, xRs, , , True
    End If
    
    If fSelVarios = True Then xRs.MoveFirst
    
    Do While Not xRs.EOF
        With fg(Index)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = xRs.Fields("idcuenta") & ""
            .TextMatrix(.Rows - 1, 2) = xRs.Fields("cuenta") & ""
            .TextMatrix(.Rows - 1, 3) = xRs.Fields("descripcion") & ""
        End With
        If fSelVarios = False Then Exit Do
        xRs.MoveNext
    Loop

SALIR:
    Agregando = False
    Set xRs = Nothing
    Exit Sub
    
error:
    Agregando = False
    Set xRs = Nothing:
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN FILA DEL CONTROL Fg
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Index     |  Integer    |  INDICA EL INDICE DEL CONTROL Fg
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel(Index As Integer)
    If QueHace = 3 Then Exit Sub
    
    If fg(Index).Row <= 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una fila correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    
    ' ELIMINAR EL REGISTRO
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    If Index = 0 Or Index = 1 Then         ' ELIMINAR TODOS LOS REGISTROS
        LimpiarGrid fg(Index + 2), True, 1
        ' 1 = ACTIVO;   2 = PASIVO
        ' ELIMINAR DATOS DEL TEMPORAL
        If fg(Index).TextMatrix(fg(Index).Row, 1) <> "" Then
            TmpCta.Filter = "tipo= " + CStr(Index + 1) + " AND orden = " + fg(Index).TextMatrix(fg(Index).Row, 1)
            If TmpCta.RecordCount <> 0 Then
                TmpCta.MoveFirst
                Do While Not TmpCta.EOF
                    TmpCta.Delete
                    TmpCta.MoveNext
                Loop
            End If
        End If
        lbl_cabecera(Index) = ""
    Else
        ' ELIMINAR SOLO UN REGISTRO
        If fg(Index).TextMatrix(fg(Index).Row, 1) <> "" Then
            If TmpCta.RecordCount <> 0 Then TmpCta.MoveFirst
            Do While Not TmpCta.EOF
                If NulosN(TmpCta.Fields("idcuenta")) = NulosN(fg(Index).TextMatrix(fg(Index).Row, 1)) Then
                    TmpCta.Delete
                End If
                TmpCta.MoveNext
            Loop
        End If
    End If
    
    fg(Index).RemoveItem (fg(Index).Row)
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        '--ACTIVO
        Case 0 '--ADD REG
            pRegistroAdd 0
        Case 1 '--DEL REG
            pRegistroDel 0
        '--DE LAS CUENTAS
        Case 4 '--ADD
            pRegistroAdd 2
        Case 5 '--SEL
            pRegistroAdd 2, True
        Case 6 '--DEL
            pRegistroDel 2
            
        '--PASIVO
        Case 2 '--ADD REG
            pRegistroAdd 1
        Case 3 '--DEL REG
            pRegistroDel 1
        '--DE LAS CUENTAS
        Case 7 '--ADD
            pRegistroAdd 3
        Case 8 '--SEL
            pRegistroAdd 3, True
        Case 9 '--DEL
            pRegistroDel 3
    End Select
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    
    If Index = 2 Or Index = 3 Then
        If fg(Index).Col = 1 Then
            fg(Index).Editable = flexEDNone
        Else
            fg(Index).Editable = flexEDKbdMouse
        End If
    Else
        If fg(Index).Col < 3 Then
            fg(Index).Editable = flexEDNone
        Else
            fg(Index).Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If Index = 2 Or Index = 3 Then
        Select Case Col
            Case 4
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
            Case Else
                KeyAscii = 0
        End Select
    Else
        Select Case Col
            Case 3
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
            Case 7
                KeyAscii = 0
        End Select
    End If
End Sub

Private Sub fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        pRegistroAdd Index
    End If
    
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        pRegistroDel Index  'F4 = Eliminar Item
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    
    If Row = 0 Then Exit Sub
    
    If Index = 0 Or Index = 1 Then
        Select Case Col
            Case 3   ' VALIDAR QUE EL NUMERO DE ORDEN SEA UNICO
                If IsNumeric(fg(Index).TextMatrix(Row, Col)) = False Then
                    MsgBox "El valor ingresado no es numérico", vbExclamation, xTitulo
                    fg(Index).TextMatrix(Row, Col) = "":    Exit Sub
                End If
                
                If GRID_BUSCAR_VALOR(fg(Index), CInt(Col), fg(Index).TextMatrix(Row, Col), False, , Row) <> "-1" Then
                    MsgBox "Se le recuerda que ya existe el número de orden" + vbCr + "Se recomienda que el número de orden sea diferente", vbInformation, xTitulo
                End If
                fg(Index).TextMatrix(Row, Col) = CInt(fg(Index).TextMatrix(Row, Col))
        End Select
    Else
        If Col <> 4 Then Exit Sub
        
        If IsNumeric(fg(Index).TextMatrix(Row, Col)) = False Then
            MsgBox "El valor ingresado no es numérico", vbExclamation, xTitulo
            fg(Index).TextMatrix(Row, Col) = "":    Exit Sub
        End If
        
        If TmpCta.State = 0 Then Exit Sub
        
        If TmpCta.EOF = False Or TmpCta.BOF = False Or TmpCta.RecordCount <> 0 Then TmpCta.MoveFirst
        
        Do While Not TmpCta.EOF
            If NulosN(fg(Index).TextMatrix(Row, 1)) = NulosN(TmpCta.Fields("idcuenta")) Then
                TmpCta.Fields("idgru") = fg(Index).TextMatrix(Row, Col)
            End If
            TmpCta.MoveNext
        Loop
    End If
    
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg_CellChanged (" + CStr(Index) + ")"
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Col <> 7 Or Index = 2 Or Index = 3 Then Exit Sub
    
    Agregando = True
    
    With FrmManFormula
        .RECIBE_LINK_FRM fg(Index), fg(Index), Row, 7, fg(Index).TextMatrix(Row, 7), "- " + IIf(Index = 0, "Activo", "Pasivo"), "2", "4"
        .Show 1
    End With
    
    Agregando = False
    Exit Sub

SALIR:
    Agregando = False
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            Select Case Index
                Case 0: PopupMenu Menu1
                
                Case 1: PopupMenu Menu2
                
                Case 2: PopupMenu Menu3
                
                Case 3: PopupMenu Menu4
            End Select
        End If
    End If
End Sub

Private Sub Fg_RowColChange(Index As Integer)
    On Error GoTo error
    
    If Agregando = True Then Exit Sub
    
    If Index = 2 Or Index = 3 Then Exit Sub
    
    If fg(Index).Rows = 1 Then
        Exit Sub
    End If
    
    fg(Index + 2).Rows = 1
    ' NOMBRE DE LA CABECERA
    lbl_cabecera(Index) = fg(Index).TextMatrix(fg(Index).Row, 4)
    
    If fg(Index).Row <= 0 Then Exit Sub
    
    If fg(Index).TextMatrix(fg(Index).Row, 2) = "" Then Exit Sub
    ' FILTRANDO LAS CUENTAS POR CABECERA
    TmpCta.Filter = "tipo = " + CStr(Index + 1) + " AND orden=" + fg(Index).TextMatrix(fg(Index).Row, 1)
    
    If TmpCta.RecordCount <> 0 Then TmpCta.MoveFirst
    Agregando = True
    
    Do While Not TmpCta.EOF
        fg(Index + 2).AddItem ""
        fg(Index + 2).TextMatrix(fg(Index + 2).Rows - 1, 1) = TmpCta.Fields("idcuenta") & ""
        fg(Index + 2).TextMatrix(fg(Index + 2).Rows - 1, 2) = TmpCta.Fields("cuenta") & ""
        fg(Index + 2).TextMatrix(fg(Index + 2).Rows - 1, 3) = TmpCta.Fields("descripcion") & ""
        fg(Index + 2).TextMatrix(fg(Index + 2).Rows - 1, 4) = TmpCta.Fields("idgru") & ""
        TmpCta.MoveNext
    Loop
    Agregando = False
    Exit Sub

error:
    Agregando = False
    SHOW_ERROR Me.Name, "Fg_RowColChange (" + CStr(Index) + ")"
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = True Then Exit Sub
    Dim Rpta As Integer

    SeEjecuto = False
    CARGAR_GRID
    SeEjecuto = True
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ninguna cuenta por rendir, ¿Desea agergar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : CARGAR_GRID
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS REGISTRO DE LA TABLA con_balance
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CARGAR_GRID()
    Dim xSql  As String
        
    xSql = "SELECT con_balance.id, con_balance.descripcion, con_balance.observacion FROM con_balance;"

    ' CARGANDO_DATOS
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, xSql, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    
    Dg3.BatchUpdates = False
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Habilitar_Obj False
    
    fg(0).Tag = fg(0).FormatString
    fg(1).Tag = fg(1).FormatString
    fg(2).Tag = fg(2).FormatString
    fg(3).Tag = fg(3).FormatString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg3.DataSource = Nothing
End Sub

Private Sub Menu1_1_Click()
    cmd_Click 0
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 1
End Sub

Private Sub Menu2_1_Click()
    cmd_Click 2
End Sub

Private Sub Menu2_3_Click()
    cmd_Click 3
End Sub

Private Sub Menu3_1_Click()
    cmd_Click 4
End Sub

Private Sub Menu3_2_Click()
    cmd_Click 5
End Sub

Private Sub Menu3_4_Click()
    cmd_Click 6
End Sub

Private Sub Menu3_5_Click()
    Dim Q_ROW As Long
    
    If fg(2).Rows <= 1 Then Exit Sub
    
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    
    Do While fg(2).Rows > 1
        fg(2).Row = 1
        cmd_Click 6
    Loop
    Agregando = False
End Sub

Private Sub Menu4_1_Click()
    cmd_Click 7
End Sub

Private Sub Menu4_2_Click()
    cmd_Click 8
End Sub

Private Sub Menu4_4_Click()
    cmd_Click 9
End Sub

Private Sub Menu4_5_Click()
    Dim Q_ROW As Long
    
    If fg(3).Rows <= 1 Then Exit Sub
    
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    
    Do While fg(3).Rows > 1
        fg(3).Row = 1
        cmd_Click 9
    Loop
    Agregando = False
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
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then RstFrm.Filter = ""
    
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA con_balance
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETE * FROM con_balancedet WHERE idcab = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_balancecab   WHERE idcab = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM con_balance WHERE id = " & RstFrm("id") & ""
        
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningún balance, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                Nuevo
            Else
                TabOne1.CurrTab = 0
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
    Label1.Caption = "Detalle del Balance"
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    Dg3.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
    
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    Habilitar_Obj True
    IdFila = 999
    IDBALANCE_CABECERA = HallaCodigoTabla("con_balancecab", xCon, "id")
    GRID_COMBOLIST fg(0), 7
    GRID_COMBOLIST fg(1), 7
    Label1.Caption = "Modificando Balance"
    txt(1).SetFocus
    ' COMODIN PARA IR INCREMENTANDO LOS REGISTROS DE LA CABECERA DEL BALANCE
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EN FORMA DETALLADA LOS DATOS DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE
'*                    DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    On Error GoTo error
    
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        
        txt(0).Text = .Fields("id") & ""               ' CODIGO
        txt(1).Text = .Fields("descripcion") & ""
        txt(2).Text = .Fields("observacion") & ""
        MuestraDetalle
    End With
    
    Exit Sub

error:
    SHOW_ERROR
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim xCol, xFil As Integer
    Dim xSql As String
    Dim xFch As Date
    Dim xFila  As Integer
    Dim XGrid As Integer
    On Error GoTo error
    
    xSql = "SELECT con_balancecab.id, con_balancecab.orden, con_balancecab.descripcion, con_balancecab.tipo, con_balancecab.negrita, con_balancecab.sallin, con_balancecab.formula " _
        + vbCr + " FROM con_balancecab " _
        + vbCr + " WHERE (((con_balancecab.idcab)=" + CStr(RstFrm.Fields("id")) + ")) " _
        + vbCr + " ORDER BY con_balancecab.tipo, con_balancecab.orden;"
        
    RST_Busq xRs, xSql, xCon
    If xRs.RecordCount <> 0 Then
        Agregando = True
        For XGrid = 0 To 1          '0 = ACTIVO  1 = PASIVO
            xRs.Filter = "tipo=" + CStr(XGrid + 1)
            If xRs.RecordCount <> 0 Then
                fg(XGrid).Rows = 1
                xRs.MoveFirst
                Do While Not xRs.EOF
                    With fg(XGrid)
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 1) = xRs.Fields("orden") & ""
                        .TextMatrix(.Rows - 1, 2) = xRs.Fields("id") & ""
                        .TextMatrix(.Rows - 1, 3) = xRs.Fields("orden") & ""
                        .TextMatrix(.Rows - 1, 4) = xRs.Fields("descripcion") & ""
                        .TextMatrix(.Rows - 1, 5) = NulosN(xRs.Fields("negrita"))
                        .TextMatrix(.Rows - 1, 6) = NulosN(xRs.Fields("sallin"))
                        .TextMatrix(.Rows - 1, 7) = xRs.Fields("formula") & ""
                        xRs.MoveNext
                    End With
                Loop
            End If
        Next XGrid
    End If
    
    ' CARGANDO DATOS DE LAS CUENTAS
    Dim N_SQL As String
    Dim RST_TMP As New ADODB.Recordset
    
    N_SQL = fGenerarConsulta(RstFrm.Fields("id"))
    If N_SQL <> "" Then
        RST_Busq RST_TMP, N_SQL, xCon
        CARGAR_RST_TMP TmpCta, RST_TMP
    End If
    
    Set RST_TMP = Nothing
    Set xRs = Nothing
    Agregando = False
    
    ' CARGANDO LOS DATOS DE LAS CUENTAS AL ACTIVO Y PASIVO
    Fg_RowColChange 0
    Fg_RowColChange 1
    Exit Sub

error:
    Set xRs = Nothing:  Set RST_TMP = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "MuestraDetalle"
End Sub

'*****************************************************************************************************
'* Nombre           : Habilitar_Obj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean    |  INDICA SI SE ACTIVA O DESACTIVA LOS CONTROLES
'* Devuelve         :
'*****************************************************************************************************
Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked txt, Not band
    habilitar cmd, band
    
    TabOne1.CurrTab = IIf(band = False, 0, 1)
    TabOne1.TabEnabled(0) = Not band
    
    If band = False Then
        fg(0).SelectionMode = flexSelectionByRow
        fg(1).SelectionMode = flexSelectionByRow
        fg(2).SelectionMode = flexSelectionByRow
        fg(3).SelectionMode = flexSelectionByRow
    Else
        fg(2).SelectionMode = flexSelectionFree
        fg(3).SelectionMode = flexSelectionFree
        fg(2).Editable = flexEDKbdMouse
        fg(3).Editable = flexEDKbdMouse
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : INICIALIZA LOS CONTROLS PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Blanquea()
    LimpiaText txt
    LimpiaText lbl_cabecera
    
    LimpiarGrid fg(0), True, 1
    LimpiarGrid fg(1), True, 1
    LimpiarGrid fg(2), True, 1
    LimpiarGrid fg(3), True, 1
    
    OCULTAR_COL fg(0), 1, 1
    OCULTAR_COL fg(1), 1, 1
    OCULTAR_COL fg(2), 1, 1
    OCULTAR_COL fg(3), 1, 1

    pDefinirRst
    TabOne2.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Nuevo()
    Dim XGrid As Integer
    
    QueHace = 1
    ActivaTool
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregando Balance"
    
    For XGrid = 0 To fg.Count - 1
        fg(XGrid).Editable = flexEDKbdMouse
        fg(XGrid).SelectionMode = flexSelectionFree
    Next XGrid
    
    TabOne2.CurrTab = 0
    txt(1).SetFocus
    
    GRID_COMBOLIST fg(0), 7
    GRID_COMBOLIST fg(1), 7
    IDBALANCE_CABECERA = HallaCodigoTabla("con_balancecab", xCon, "id")
    IdFila = 1
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA con_balance, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " El Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstCta As New ADODB.Recordset
    Dim TmpRst As New ADODB.Recordset        ' PARA BUSCAR SI EL NUMERO DE CABECERA YA ESTA REGISTRADO
    
    Dim xCod As Integer
    Dim xCodDet As Integer '--al detalle
    Dim XGrid As Integer
    Dim xCol, xFil As Integer
    Dim xCorr As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO OBTENEMOS EL ULTIMO ID DE LA TABLA con_balance
        xCod = HallaCodigoTabla("con_balance", xCon, "id")
        RST_Busq RstCab, "SELECT top 1 * FROM con_balance ", xCon
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        ' SI SE ESTA MODIFICANDO UN REGISTRO OBTENEMOS EL ID DEL REGISTRO ACTUAL
        xCod = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_balance WHERE id =" & RstFrm("id") & "", xCon
        xCon.Execute "DELETE * FROM con_balancedet WHERE idcab = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_balancecab   WHERE idcab = " & RstFrm("id") & ""
    End If
    
    RST_Busq RstDet, "SELECT top 1 * FROM con_balancecab", xCon
    RST_Busq RstCta, "SELECT top 1 * FROM con_balancedet", xCon
    
    RstCab("descripcion") = Trim(txt(1).Text) & ""
    RstCab("observacion") = Trim(txt(2).Text) & ""
    
    RstCab.Update
    
    ' GRABAMOS LOS DETALLES
    For XGrid = 0 To 1
        With fg(XGrid)
            For xFil = 1 To .Rows - 1
                If NulosN(.TextMatrix(xFil, 1)) > 0 And .TextMatrix(xFil, 3) <> "" Then
                    RstDet.AddNew
                    ' LLAVE
                    xCodDet = NulosN(.TextMatrix(xFil, 2))
                    If QueHace = 1 Then ' NUEVO
                        xCodDet = HallaCodigoTabla("con_balancecab", xCon, "id")
                        ' VERIFICAR SI SE REPITE SI YA EXISTE EL NUMERO
                        RST_Busq TmpRst, "SELECT con_balancecab.id FROM con_balancecab WHERE (((con_balancecab.id)=" + CStr(xCodDet) + "));", xCon
                        If TmpRst.EOF = False Or TmpRst.BOF = False Or TmpRst.RecordCount <> 0 Then
                            xCodDet = HallaCodigoTabla("con_balancecab", xCon, "id")
                            ' TENEMOS QUE ACTUALIZAR DATOS DE LA FORMULA
                        End If
                        Set TmpRst = Nothing
                    End If
                    RstDet("idcab") = xCod
                    RstDet("id") = xCodDet
                    ' FIN
                    RstDet("descripcion") = .TextMatrix(xFil, 4)
                    RstDet("tipo") = CStr(XGrid + 1)
                    RstDet("orden") = NulosN(.TextMatrix(xFil, 3))
                    RstDet("negrita") = Val(.TextMatrix(xFil, 5))
                    RstDet("sallin") = Val(.TextMatrix(xFil, 6))
                    RstDet("formula") = .TextMatrix(xFil, 7)
                    RstDet.Update
                    
                    TmpCta.Filter = "tipo = " + CStr(XGrid + 1) + " AND orden=" + .TextMatrix(xFil, 1)
                    If TmpCta.RecordCount > 0 Then
                        TmpCta.MoveFirst
                        Do While Not TmpCta.EOF
                            RstCta.AddNew
                            ' CLAVE
                            RstCta("idcab") = xCod
                            RstCta("idbal") = xCodDet
                            ' FIN CLAVE
                            RstCta("idcuenta") = TmpCta.Fields("idcuenta")
                            RstCta("idgru") = TmpCta.Fields("idgru")
                            RstCta.Update
                            TmpCta.MoveNext
                        Loop
                    End If
                End If
            Next xFil
        End With
    Next XGrid
    
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    xCon.CommitTrans
    Grabar = True
SALIR:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstCta = Nothing:    Set TmpRst = Nothing
    Me.MousePointer = vbDefault
    Exit Function

LaCague:
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstCta = Nothing:    Set TmpRst = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VALIDA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    ' VALIDAR QUE LA GRILLA DE ACTIVO Y PASIVO TENGAN VALORES TANTO DE ORDEN Y DESCRIPCION
    Dim Q_ROW  As Long
    Dim QGrid As Integer
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "Ingrese la descripción del balance", vbExclamation, xTitulo
        Exit Function
    End If
    
    ' VALIDAR QUE EL REGISTRO NO ESTE REGISTRADO
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq RstTmp, "SELECT descripcion FROM con_balance WHERE ucase(descripcion)='" + UCase(Trim(txt(1).Text)) + "';", xCon
    Else
        RST_Busq RstTmp, "SELECT descripcion FROM con_balance WHERE ucase(descripcion)='" + UCase(Trim(txt(1).Text)) + "' AND id <> " + CStr(RstFrm.Fields("id")) + ";", xCon
    End If
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "El registro " + IIf(QueHace = 1, " ya fue ingresado", "ya existe"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    Set RstTmp = Nothing
    
    For QGrid = 0 To 1
        With fg(QGrid)
            For Q_ROW = 1 To .Rows - 1
                If IsNumeric(.TextMatrix(Q_ROW, 3)) = False Or .TextMatrix(Q_ROW, 3) = "0" Then
                    MsgBox "Ingrese El N° de Orden:", vbExclamation, xTitulo
                    TabOne2.CurrTab = QGrid
                    Agregando = True:  .Row = Q_ROW: .Col = 3: Agregando = False
                    
                    Exit Function
                ElseIf .TextMatrix(Q_ROW, 4) = "" Then
                    MsgBox "Ingrese la Descripción:", vbExclamation, xTitulo
                    TabOne2.CurrTab = QGrid
                    Agregando = True:  .Row = Q_ROW: .Col = 4: Agregando = False
                    
                    Exit Function
                End If
            Next Q_ROW
        End With
    Next QGrid
    fValidarDatos = True
End Function
 
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Imprimir True

    If ButtonMenu.Index = 2 Then Imprimir
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim N_SQL As String
   
    Dim xCampos(1, 4) As String
    
    xCampos(0, 0) = "Descripción":        xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":     xCampos(0, 3) = "C"
        
    N_SQL = "SELECT con_balance.id, con_balance.descripcion, con_balance.observacion FROM con_balance;"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Balance", "descripcion", "descripcion", Principio
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
'* Descripcion      : EJECUTA UN FILTRO EN EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Filtrar()
    Dim xCampos(0, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "Descripcion":   xCampos(0, 2) = "C":     xCampos(0, 3) = "1000"
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Imprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL BALANCE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Imprimir(Optional IMP_LISTADO As Boolean = False)
    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE BALANCE", " "
    End If

    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "IMPRIMIR"
End Sub

Private Function fGenerarConsulta(X_ID As String) As String
    Dim N_SQL As String

    N_SQL = "SELECT con_balancecab.tipo,con_balancecab.orden, con_balancedet.idbal, con_balancedet.idcuenta, con_planctas.cuenta, con_planctas.descripcion, con_balancedet.idgru " _
        + vbCr + " FROM con_balancecab INNER JOIN (con_planctas INNER JOIN con_balancedet ON con_planctas.id = con_balancedet.idcuenta) ON (con_balancecab.id = con_balancedet.idbal) AND (con_balancecab.idcab = con_balancedet.idcab) " _
        + vbCr + " WHERE (((con_balancedet.idcab)=" + X_ID + "));"

    fGenerarConsulta = N_SQL
End Function

'*****************************************************************************************************
'* Nombre           : pDefinirRst
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DEFINIR EL RECORSET TEMPORAL PARA INSUMO Y TAREA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pDefinirRst()
    Dim RST_ORIGEN As New ADODB.Recordset
    Dim N_SQL As String
    N_SQL = fGenerarConsulta("-1")
    RST_Busq RST_ORIGEN, N_SQL, xCon
    DEFINIR_RST_TMP TmpCta, RST_ORIGEN
    Set RST_ORIGEN = Nothing
End Sub
