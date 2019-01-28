VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmManProduccion2 
   Caption         =   "Produccion - Ingreso de Producción"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   7
      Top             =   375
      Width           =   11910
      _cx             =   21008
      _cy             =   12779
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6825
         Left            =   12555
         TabIndex        =   11
         Top             =   375
         Width           =   11820
         Begin VB.Frame Frame4 
            Height          =   660
            Left            =   5340
            TabIndex        =   17
            Top             =   360
            Width           =   4905
            Begin VB.CommandButton cmd 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   420
               Index           =   1
               Left            =   3210
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "Eliminar Documentos Seleccionados"
               Top             =   180
               Width           =   1515
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar de Programación"
               Enabled         =   0   'False
               Height          =   420
               Index           =   3
               Left            =   5310
               TabIndex        =   29
               TabStop         =   0   'False
               ToolTipText     =   "Agregar Documentos"
               Top             =   180
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Seleccionar Producto"
               Enabled         =   0   'False
               Height          =   420
               Index           =   2
               Left            =   1710
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "Agregar Documentos"
               Top             =   180
               Width           =   1515
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Producto"
               Enabled         =   0   'False
               Height          =   420
               Index           =   0
               Left            =   210
               TabIndex        =   2
               ToolTipText     =   "Agregar Documentos"
               Top             =   180
               Width           =   1515
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   660
            Left            =   10350
            TabIndex        =   30
            Top             =   360
            Width           =   1410
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo"
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
               Index           =   1
               Left            =   90
               TabIndex        =   31
               Top             =   240
               Width           =   1245
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   2760
            TabIndex        =   22
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   345
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   1860
            Picture         =   "FrmManProduccion2.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   720
            Width           =   225
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   3060
            Left            =   45
            TabIndex        =   14
            Top             =   3735
            Width           =   11700
            _cx             =   20637
            _cy             =   5397
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
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "   Insumos  |   Tareas   "
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
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   2640
               Left            =   -12255
               TabIndex        =   16
               Top             =   45
               Width           =   11610
               Begin TrueOleDBGrid70.TDBGrid Dg 
                  Height          =   2490
                  Index           =   0
                  Left            =   75
                  TabIndex        =   4
                  Top             =   105
                  Width           =   11445
                  _ExtentX        =   20188
                  _ExtentY        =   4392
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "IDPARTE"
                  Columns(0).DataField=   "idparte"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "IDREC"
                  Columns(1).DataField=   "idrec"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "IDITEM"
                  Columns(2).DataField=   "iditem"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   0
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "Tipo Producto"
                  Columns(3).DataField=   "tipprodesc"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(4)._VlistStyle=   0
                  Columns(4)._MaxComboItems=   5
                  Columns(4).Caption=   "Insumo"
                  Columns(4).DataField=   "descripcion"
                  Columns(4).ButtonPicture.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
                  Columns(4).ButtonPicture(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
                  Columns(4).ButtonPicture(1)=   "AAAAAADAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(2)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(3)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(4)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
                  Columns(4).ButtonPicture(5)=   "AAAAAAAAAAAAAADAwMDAwMDAwMDAwMD///+AgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAA"
                  Columns(4).ButtonPicture(6)=   "AADAwMDAwMDAwMDAwMD////AwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMCAgIAAAADAwMDAwMDA"
                  Columns(4).ButtonPicture(7)=   "wMDAwMD////AwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMCAgIAAAADAwMDAwMDAwMDAwMD////A"
                  Columns(4).ButtonPicture(8)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMCAgIAAAADAwMDAwMDAwMDAwMD////AwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(9)=   "wMDAwMDAwMDAwMDAwMDAwMCAgIAAAADAwMDAwMDAwMDAwMD////AwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(10)=   "wMDAwMDAwMCAgIAAAADAwMDAwMDAwMDAwMD/////////////////////////////////////////"
                  Columns(4).ButtonPicture(11)=   "//8AAADAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(12)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(13)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA"
                  Columns(4).ButtonPicture(14)=   "wMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMA="
                  Columns(4).ButtonPicture.vt=   9
                  Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(5)._VlistStyle=   0
                  Columns(5)._MaxComboItems=   5
                  Columns(5).Caption=   "U.M."
                  Columns(5).DataField=   "abrev"
                  Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(6)._VlistStyle=   0
                  Columns(6)._MaxComboItems=   5
                  Columns(6).Caption=   "Unid."
                  Columns(6).DataField=   "unid"
                  Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(7)._VlistStyle=   0
                  Columns(7)._MaxComboItems=   5
                  Columns(7).Caption=   "Cant. Prog."
                  Columns(7).DataField=   "canprog"
                  Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(8)._VlistStyle=   0
                  Columns(8)._MaxComboItems=   5
                  Columns(8).Caption=   "Cant. Teor."
                  Columns(8).DataField=   "canteo"
                  Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(9)._VlistStyle=   0
                  Columns(9)._MaxComboItems=   5
                  Columns(9).Caption=   "Cant. Real"
                  Columns(9).DataField=   "canreal"
                  Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(10)._VlistStyle=   0
                  Columns(10)._MaxComboItems=   5
                  Columns(10).Caption=   "Diferencia"
                  Columns(10).DataField=   "dif"
                  Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   11
                  Splits(0)._UserFlags=   0
                  Splits(0).Locked=   -1  'True
                  Splits(0).MarqueeStyle=   3
                  Splits(0).AllowSizing=   -1  'True
                  Splits(0).RecordSelectorWidth=   265
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).DividerColor=   12632256
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=11"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
                  Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(8)=   "Column(1).Width=900"
                  Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=820"
                  Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8705"
                  Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
                  Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(15)=   "Column(2).Width=794"
                  Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=714"
                  Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=8705"
                  Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
                  Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(22)=   "Column(3).Width=2805"
                  Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2725"
                  Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
                  Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=8448"
                  Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
                  Splits(0)._ColumnProps(28)=   "Column(4).Width=8625"
                  Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
                  Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=8546"
                  Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
                  Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=8704"
                  Splits(0)._ColumnProps(33)=   "Column(4).Button=1"
                  Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
                  Splits(0)._ColumnProps(35)=   "Column(5).Width=1217"
                  Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
                  Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=1138"
                  Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
                  Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=8448"
                  Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
                  Splits(0)._ColumnProps(41)=   "Column(6).Width=1535"
                  Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
                  Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=1455"
                  Splits(0)._ColumnProps(44)=   "Column(6)._EditAlways=0"
                  Splits(0)._ColumnProps(45)=   "Column(6).AllowSizing=0"
                  Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=8962"
                  Splits(0)._ColumnProps(47)=   "Column(6).Visible=0"
                  Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
                  Splits(0)._ColumnProps(49)=   "Column(7).Width=1958"
                  Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
                  Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=1879"
                  Splits(0)._ColumnProps(52)=   "Column(7)._EditAlways=0"
                  Splits(0)._ColumnProps(53)=   "Column(7).AllowSizing=0"
                  Splits(0)._ColumnProps(54)=   "Column(7)._ColStyle=514"
                  Splits(0)._ColumnProps(55)=   "Column(7).Visible=0"
                  Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
                  Splits(0)._ColumnProps(57)=   "Column(8).Width=2196"
                  Splits(0)._ColumnProps(58)=   "Column(8).DividerColor=0"
                  Splits(0)._ColumnProps(59)=   "Column(8)._WidthInPix=2117"
                  Splits(0)._ColumnProps(60)=   "Column(8)._EditAlways=0"
                  Splits(0)._ColumnProps(61)=   "Column(8)._ColStyle=8962"
                  Splits(0)._ColumnProps(62)=   "Column(8).Order=9"
                  Splits(0)._ColumnProps(63)=   "Column(9).Width=2143"
                  Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
                  Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=2064"
                  Splits(0)._ColumnProps(66)=   "Column(9)._EditAlways=0"
                  Splits(0)._ColumnProps(67)=   "Column(9)._ColStyle=770"
                  Splits(0)._ColumnProps(68)=   "Column(9).Order=10"
                  Splits(0)._ColumnProps(69)=   "Column(10).Width=2011"
                  Splits(0)._ColumnProps(70)=   "Column(10).DividerColor=0"
                  Splits(0)._ColumnProps(71)=   "Column(10)._WidthInPix=1931"
                  Splits(0)._ColumnProps(72)=   "Column(10)._EditAlways=0"
                  Splits(0)._ColumnProps(73)=   "Column(10).AllowSizing=0"
                  Splits(0)._ColumnProps(74)=   "Column(10)._ColStyle=8962"
                  Splits(0)._ColumnProps(75)=   "Column(10).Visible=0"
                  Splits(0)._ColumnProps(76)=   "Column(10).Order=11"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   0
                  ColumnFooters   =   -1  'True
                  DefColWidth     =   0
                  HeadLines       =   1
                  FootLines       =   1
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  DeadAreaBackColor=   12632256
                  RowDividerColor =   12632256
                  RowSubDividerColor=   12632256
                  DirectionAfterEnter=   0
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
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=2"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H400000&"
                  _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
                  _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&,.fgcolor=&H800000&"
                  _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&H8000000F&"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&HFFFFFF&"
                  _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
                  _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.alignment=2,.locked=-1"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2,.locked=-1"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
                  _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14,.alignment=0"
                  _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
                  _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
                  _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
                  _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
                  _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
                  _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14,.alignment=0"
                  _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
                  _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
                  _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14,.alignment=1"
                  _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
                  _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
                  _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
                  _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
                  _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
                  _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
                  _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=90,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=87,.parent=14,.alignment=1"
                  _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=88,.parent=15"
                  _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=89,.parent=17"
                  _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HC0C0FF&"
                  _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14,.alignment=1"
                  _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
                  _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
                  _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=58,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14,.alignment=1"
                  _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
                  _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
                  _StyleDefs(80)  =   "Named:id=33:Normal"
                  _StyleDefs(81)  =   ":id=33,.parent=0"
                  _StyleDefs(82)  =   "Named:id=34:Heading"
                  _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(84)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(85)  =   "Named:id=35:Footing"
                  _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(87)  =   "Named:id=36:Selected"
                  _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(89)  =   "Named:id=37:Caption"
                  _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(91)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(93)  =   "Named:id=39:EvenRow"
                  _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(95)  =   "Named:id=40:OddRow"
                  _StyleDefs(96)  =   ":id=40,.parent=33"
                  _StyleDefs(97)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(98)  =   ":id=41,.parent=34"
                  _StyleDefs(99)  =   "Named:id=42:FilterBar"
                  _StyleDefs(100) =   ":id=42,.parent=33"
               End
            End
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   2640
               Left            =   45
               TabIndex        =   15
               Top             =   45
               Width           =   11610
               Begin TrueOleDBGrid70.TDBGrid Dg 
                  Height          =   2490
                  Index           =   1
                  Left            =   75
                  TabIndex        =   5
                  Top             =   105
                  Width           =   11445
                  _ExtentX        =   20188
                  _ExtentY        =   4392
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "IDPARTE"
                  Columns(0).DataField=   "idparte"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "IDREC"
                  Columns(1).DataField=   "idrec"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "IDTAR"
                  Columns(2).DataField=   "idtar"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   0
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "Tarea"
                  Columns(3).DataField=   "descripcion"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(4)._VlistStyle=   0
                  Columns(4)._MaxComboItems=   5
                  Columns(4).Caption=   "U.M."
                  Columns(4).DataField=   "abrev"
                  Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(5)._VlistStyle=   0
                  Columns(5)._MaxComboItems=   5
                  Columns(5).Caption=   "Unid."
                  Columns(5).DataField=   "unid"
                  Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(6)._VlistStyle=   0
                  Columns(6)._MaxComboItems=   5
                  Columns(6).Caption=   "Hora Inicio"
                  Columns(6).DataField=   "horini"
                  Columns(6).EditMask=   "##:##"
                  Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(7)._VlistStyle=   0
                  Columns(7)._MaxComboItems=   5
                  Columns(7).Caption=   "Hora Fin"
                  Columns(7).DataField=   "horfin"
                  Columns(7).EditMask=   "##:##"
                  Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(8)._VlistStyle=   0
                  Columns(8)._MaxComboItems=   5
                  Columns(8).Caption=   "Cant.Personas"
                  Columns(8).DataField=   "canper"
                  Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   9
                  Splits(0)._UserFlags=   0
                  Splits(0).Locked=   -1  'True
                  Splits(0).MarqueeStyle=   3
                  Splits(0).RecordSelectorWidth=   265
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).DividerColor=   12632256
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=9"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
                  Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(8)=   "Column(1).Width=1270"
                  Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1191"
                  Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8708"
                  Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
                  Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(15)=   "Column(2).Width=1191"
                  Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=1111"
                  Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=8705"
                  Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
                  Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(22)=   "Column(3).Width=9155"
                  Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=9075"
                  Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
                  Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=8704"
                  Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
                  Splits(0)._ColumnProps(28)=   "Column(4).Width=1270"
                  Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
                  Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1191"
                  Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
                  Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=8448"
                  Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
                  Splits(0)._ColumnProps(34)=   "Column(5).Width=2249"
                  Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
                  Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2170"
                  Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
                  Splits(0)._ColumnProps(38)=   "Column(5).AllowSizing=0"
                  Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=8962"
                  Splits(0)._ColumnProps(40)=   "Column(5).Visible=0"
                  Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
                  Splits(0)._ColumnProps(42)=   "Column(6).Width=2170"
                  Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
                  Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=2090"
                  Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
                  Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=513"
                  Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
                  Splits(0)._ColumnProps(48)=   "Column(7).Width=2275"
                  Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
                  Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=2196"
                  Splits(0)._ColumnProps(51)=   "Column(7)._EditAlways=0"
                  Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=513"
                  Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
                  Splits(0)._ColumnProps(54)=   "Column(8).Width=2566"
                  Splits(0)._ColumnProps(55)=   "Column(8).DividerColor=0"
                  Splits(0)._ColumnProps(56)=   "Column(8)._WidthInPix=2487"
                  Splits(0)._ColumnProps(57)=   "Column(8)._EditAlways=0"
                  Splits(0)._ColumnProps(58)=   "Column(8)._ColStyle=514"
                  Splits(0)._ColumnProps(59)=   "Column(8).Order=9"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   0
                  ColumnFooters   =   -1  'True
                  DefColWidth     =   0
                  HeadLines       =   1
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
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H400000&"
                  _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2,.locked=-1"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
                  _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
                  _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
                  _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
                  _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=86,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14,.alignment=0"
                  _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
                  _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
                  _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14,.alignment=1"
                  _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
                  _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
                  _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2,.bgcolor=&H9B9BFF&"
                  _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
                  _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
                  _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
                  _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2,.bgcolor=&HAEAEFF&"
                  _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
                  _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
                  _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
                  _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1,.bgcolor=&HC0C0FF&"
                  _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
                  _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
                  _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
                  _StyleDefs(72)  =   "Named:id=33:Normal"
                  _StyleDefs(73)  =   ":id=33,.parent=0"
                  _StyleDefs(74)  =   "Named:id=34:Heading"
                  _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(76)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(77)  =   "Named:id=35:Footing"
                  _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(79)  =   "Named:id=36:Selected"
                  _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(81)  =   "Named:id=37:Caption"
                  _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(83)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(85)  =   "Named:id=39:EvenRow"
                  _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(87)  =   "Named:id=40:OddRow"
                  _StyleDefs(88)  =   ":id=40,.parent=33"
                  _StyleDefs(89)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(90)  =   ":id=41,.parent=34"
                  _StyleDefs(91)  =   "Named:id=42:FilterBar"
                  _StyleDefs(92)  =   ":id=42,.parent=33"
               End
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   0
            Left            =   900
            TabIndex        =   0
            Top             =   345
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
            Valor           =   "21/11/2007"
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   900
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(1)"
            ToolTipText     =   "Ingrese DNI del Supervisor"
            Top             =   690
            Width           =   1215
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2460
            Left            =   45
            TabIndex        =   3
            Top             =   1065
            Width           =   11700
            _cx             =   20637
            _cy             =   4339
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManProduccion2.frx":0132
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
            Caption         =   "Producto:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   3510
            Width           =   11610
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   2205
            TabIndex        =   23
            Top             =   465
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Supervisor"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   21
            Top             =   780
            Width           =   750
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
            Height          =   300
            Index           =   1
            Left            =   2130
            TabIndex        =   20
            Top             =   690
            Width           =   3135
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(1)"
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
            Height          =   300
            Index           =   1
            Left            =   3975
            TabIndex        =   19
            Top             =   345
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fch Prod"
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   13
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Producción"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   75
            TabIndex        =   12
            Top             =   15
            Width           =   11610
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   45
         TabIndex        =   8
         Top             =   375
         Width           =   11820
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   45
            TabIndex        =   28
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11404
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
            Columns(1).Caption=   "Fch. Pro."
            Columns(1).DataField=   "dia1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Num. Reg."
            Columns(2).DataField=   "numparte"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Receta"
            Columns(3).DataField=   "codrec"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Producto"
            Columns(4).DataField=   "proddesc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Resp. de la Producción"
            Columns(5).DataField=   "resnom"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "UM"
            Columns(6).DataField=   "abrev"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Hor. Ini."
            Columns(7).DataField=   "horiniF"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Hor. Fin."
            Columns(8).DataField=   "horfinF"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Can. Prod."
            Columns(9).DataField=   "cantidad1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Num. Ord. Prod."
            Columns(10).DataField=   "numordprod"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Estado"
            Columns(11).DataField=   "desestado"
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
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1482"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1402"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1826"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1746"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=5636"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=5556"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=6191"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6112"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=847"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=767"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1191"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1111"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1217"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1138"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1614"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1535"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=516"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(68)=   "Column(11).Width=2011"
            Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=1931"
            Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(72)=   "Column(11)._ColStyle=516"
            Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
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
            HeadLines       =   1
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=90,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=86,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=62,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=66,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=63,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=64,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=65,.parent=17"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9705
            TabIndex        =   26
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Producción"
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
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   10
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblMes 
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   8835
            TabIndex        =   9
            Top             =   30
            Width           =   1275
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6600
         Top             =   0
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
               Picture         =   "FrmManProduccion2.frx":03A1
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":08E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":0C77
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":0DFB
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":124F
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":1367
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":18AB
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":1DEF
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":1F03
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":2017
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":246B
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":25D7
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion2.frx":2B1F
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_3 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Seleccionar Producto"
      End
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar de Programación"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Eliminar Producto"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Eliminar Producto"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "Eliminar Todo"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu Menu3_1 
         Caption         =   "Seleccionar un Producto"
      End
      Begin VB.Menu Menu3_2 
         Caption         =   "Seleccionar Varios Productos"
      End
      Begin VB.Menu Menu3_3 
         Caption         =   "Productos Con Programación"
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "Menu4"
      Visible         =   0   'False
      Begin VB.Menu Menu4_1 
         Caption         =   "Insumo"
         Begin VB.Menu Menu4_1_1 
            Caption         =   "x Producto"
         End
         Begin VB.Menu Menu4_1_2 
            Caption         =   "Todos los Productos"
         End
         Begin VB.Menu Menu4_1_3 
            Caption         =   "Resumen"
         End
      End
   End
   Begin VB.Menu Menu5 
      Caption         =   "Menu5"
      Visible         =   0   'False
      Begin VB.Menu Menu5_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu5_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu5_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmManProduccion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANPRODUCCION.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE EL INGRESO Y MODIFICACION DE LOS PARTES DE PRODUCCION
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 05/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer                                    ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim Agregando As Boolean                                  ' INDICA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID
Dim SeEjecuto As Boolean                                  ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim RstFrm As New ADODB.Recordset                         ' RECORDSET PARA ALMACENAR LOS DATOS  DE LA TABLA pro_produccion
Dim mMesActivo As Integer                                 ' INDICA EL MES ACTIVO
Dim RST_INSUMO As New ADODB.Recordset                     ' PARA LOS INSUMOS
Dim RST_TAREA As New ADODB.Recordset                      ' PARA LAS TAREAS
Dim M_NUM_PARTE As Long                                   ' INDICA EL NUMERO PARTE DE PRODUCCION
Private Const FORMAT_NUM_PRODUCCION As String = "000000"  ' INDICA EL FORMATO DE LA COLUMNA CON EL NUMERO DE PRODUCCION
Private Const FORMAT_NUM_PARTE As String = "00000000"     ' INDICA EL FORMATO DE LA COLUMNA CON EL NUMERO DE PRODUCCION
Private Const FORMAT_CANT As String = "#0.000000"
Dim fOrdenLista As Boolean                                ' especfica el orden de la lista de la consulta
Dim mIdRegistro&                                          ' identificador del registro
Dim fCierrePeriodo As Boolean                             '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

Dim COLUMNANUMPROD_ As Integer
Dim COLUMNAPRODUCTO_ As Integer
Dim COLUMNARECETA_ As Integer
Dim COLUMNAUM_ As Integer
Dim COLUMNAORDENPROD_ As Integer
Dim COLUMNATOTPROG_ As Integer
Dim COLUMNATOTPROD_ As Integer
Dim COLUMNAHORINI_ As Integer
Dim COLUMNAHORFIN_ As Integer
Dim COLUMNARESPONSABLE_ As Integer
Dim COLUMNATURNO_ As Integer
Dim COLUMNAESTADO_ As Integer
Dim COLUMNAIDPARTE_ As Integer
Dim COLUMNAIDREC_ As Integer
Dim COLUMNAIDITEM_ As Integer
Dim COLUMNAIDUNID_ As Integer
Dim COLUMNAIDRES_ As Integer
Dim COLUMNAIDTURNO_ As Integer
Dim COLUMNACORR_ As Integer
Dim COLUMNAOBS_ As Integer
Dim COLUMNAIDORD_ As Integer

Dim NUMEROPROD_ As Double
Dim cSQL As String
Dim CORR_ As Double
Dim CAMBIOGRABAR_ As Double
Dim ESTADOANTERIOR_ As Double
    
Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4


'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UNA FILA AL CONTROL Fg1
'* Parametros       : NOMBRE                    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    fSeleccionVarios          |  Boolean   |
'*                    fAddRegistroSinPrograma   |  Boolean   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True, Optional fAddRegistroSinPrograma As Boolean = False)
    Dim SQL_IDREC As String
    Dim nSQL As String
    Dim xRs  As New ADODB.Recordset
    Dim A As Integer
    Dim xFila As Integer
    Dim RstReceta As New ADODB.Recordset         ' BUSCAR LA RECETA PREDETERMINADA
    
    If IsDate(TxtFecha(0).valor) = False Then
        MsgBox "Ingrese la fecha de Producción", vbExclamation, xTitulo
        Exit Sub
    End If
        
    ' GENERAR EL WHERE DE LOS ID'S RECETA PARA QUE NO SE REPITAN
    If fAddRegistroSinPrograma = False Then SQL_IDREC = GENERAR_SQL_ID(Fg1, COLUMNAIDREC_, "pro_receta.id", "NOT IN")
    If SQL_IDREC <> "" Then SQL_IDREC = " AND " + SQL_IDREC
    
    On Error GoTo error
    
    If fAddRegistroSinPrograma = False Then
        ReDim xCampos(4, 5) As String
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
        xCampos(1, 0) = "Receta":           xCampos(1, 1) = "codrec":       xCampos(1, 2) = "1000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "S"
        xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Cant.Prog":        xCampos(3, 1) = "canprog":      xCampos(3, 2) = "1000":      xCampos(3, 3) = "N":    xCampos(3, 4) = "N"

        nSQL = "SELECT pro_programadet.idprod, pro_receta.id AS idrec, pro_receta.iditem,pro_receta.idunimed, alm_inventario.descripcion, pro_receta.codrec, mae_unidades.abrev, pro_programadet.canpro as canprog " _
            + vbCr + " FROM (alm_inventario INNER JOIN (pro_receta LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec " _
            + vbCr + " WHERE (((pro_programadet.idpro) = 0) And ((pro_programadet.dia) = CDATE('" + TxtFecha(0).valor + "'))) " + SQL_IDREC _
            + vbCr + " ORDER BY alm_inventario.descripcion;"
    Else
        ReDim xCampos(4, 5) As String
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
        xCampos(1, 0) = "Familia":          xCampos(1, 1) = "famdesc":      xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Codigo":           xCampos(3, 1) = "codpro":       xCampos(3, 2) = "1500":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"

        nSQL = "SELECT alm_inventario.id as iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_familia.descripcion AS famdesc, mae_unidades.abrev,0 AS canprog " _
            + vbCr + " FROM mae_unidades RIGHT JOIN (alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) ON mae_unidades.id = alm_inventario.idunimed " _
            + vbCr + " WHERE alm_inventario.tippro IN (3, 8)  AND alm_inventario.activo = -1 " _
            + vbCr + " ORDER BY alm_inventario.descripcion, mae_familia.descripcion;"
    End If
    
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Programas de Producción para el dia " + TxtFecha(0).valor
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Programas de Producción para el dia " + TxtFecha(0).valor, "descripcion", "codpro", Principio
    End If
    
    Agregando = True
    If xRs.State = 0 Then GoTo SALIR
    
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    If fSeleccionVarios = True Then xRs.MoveFirst
   
    ' SI NO HAY REGISTROS OBTENER EL NUMERO DE PRODUCCION
    'If Fg1.Rows = 1 Then M_NUM_PARTE = HallaValor(xCon, "pro_producciondet", "numparte")
    
    If NUMEROPROD_ = 0 Then NUMEROPROD_ = HallaValor(xCon, "pro_producciondet", "numparte")
    M_NUM_PARTE = NUMEROPROD_
    
    Do While Not xRs.EOF
        ADD_REG Fg1, Fila_Ninguno
        With Fg1
            ' DEL NUMERO DE REGISTRO
            .TextMatrix(.Rows - 1, COLUMNANUMPROD_) = M_NUM_PARTE
            .TextMatrix(.Rows - 1, COLUMNAPRODUCTO_) = xRs.Fields("descripcion") & ""
            .TextMatrix(.Rows - 1, COLUMNATOTPROG_) = Format(NulosN(xRs.Fields("canprog")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNAIDPARTE_) = M_NUM_PARTE
            .TextMatrix(.Rows - 1, COLUMNAIDITEM_) = xRs.Fields("iditem") & ""
            .TextMatrix(.Rows - 1, COLUMNACORR_) = CORR_
            CORR_ = CORR_ + 1
            NUMEROPROD_ = NUMEROPROD_ + 1
            
            If fAddRegistroSinPrograma = False Then
                ' YA FUE PROGRAMADO
                .TextMatrix(.Rows - 1, COLUMNARECETA_) = xRs.Fields("codrec") & ""
                .TextMatrix(.Rows - 1, COLUMNAUM_) = xRs.Fields("abrev") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDREC_) = xRs.Fields("idrec") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDUNID_) = xRs.Fields("idunimed") & ""

                DATOS_TMP_ADD CStr(M_NUM_PARTE), xRs.Fields("idrec"), E_INSUMO, True, fAddRegistroSinPrograma
                DATOS_TMP_ADD CStr(M_NUM_PARTE), xRs.Fields("idrec"), e_TAREA, True, fAddRegistroSinPrograma
            Else
                ' CARGAR RECETA PREDETERMINADA
                RST_Busq RstReceta, "SELECT TOP 1 pro_receta.id AS idrec, pro_receta.descripcion, pro_receta.codrec, mae_unidades.abrev , pro_receta.idunimed" _
                    + vbCr + " FROM pro_receta INNER JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id " _
                    + vbCr + " Where (((pro_receta.iditem) = " + CStr(xRs.Fields("iditem")) + "))" _
                    + vbCr + " ORDER BY pro_receta.prirec;", xCon
                
                If RstReceta.EOF = False Or RstReceta.BOF = False Or RstReceta.RecordCount <> 0 Then
                    If VERIFICAR_LISTA(Fg1, COLUMNARECETA_, RstReceta.Fields("codrec") & "", False) = True Then
                        .TextMatrix(.Rows - 1, COLUMNARECETA_) = RstReceta.Fields("codrec") & ""
                        .TextMatrix(.Rows - 1, COLUMNAUM_) = RstReceta.Fields("abrev") & ""
                        .TextMatrix(.Rows - 1, COLUMNAIDREC_) = RstReceta.Fields("idrec") & ""
                        .TextMatrix(.Rows - 1, COLUMNAIDUNID_) = RstReceta.Fields("idunimed") & ""
                        
                        DATOS_TMP_ADD CStr(M_NUM_PARTE), RstReceta.Fields("idrec"), E_INSUMO, True, fAddRegistroSinPrograma
                        DATOS_TMP_ADD CStr(M_NUM_PARTE), RstReceta.Fields("idrec"), e_TAREA, True, fAddRegistroSinPrograma
                    End If
                End If
            End If
            
            .TextMatrix(.Rows - 1, COLUMNAESTADO_) = ESTADOPROCESADO_

            If fSeleccionVarios = False Then Exit Do
        End With
        
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop

SALIR:
    Agregando = False
    If Fg1.Rows >= 2 Then Fg1.Row = Fg1.Rows - 1: Fg1.Col = COLUMNATOTPROD_:
    Set xRs = Nothing
    Fg1.SetFocus
    Exit Sub

error:
    'Resume
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

'*****************************************************************************************************
'* Nombre           : pGenerarConsulta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESTA FUNCION CONSTRUYE LAS CONSULTAS DE INSUMOS Y TAREAS EN FUNCION DE LA RECETA
'*                    (AGREGAR NUEVA RECETA, AGREGAR DE PROGRAMACION DE PRODUCCION)
'* Parametros       : NOMBRE                   |  TIPO        |   DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    mIdParte                 |  String      |
'*                    mIdReceta                |  String      |
'*                    fTipo                    |  e_PROGRAMA  |
'*                    dFecha                   |  Date        |
'*                    fAddRegistro             |  Boolean     |
'*                    fAddRegistroSinPrograma  |  Boolean     |
'* Devuelve         : String
'*****************************************************************************************************
Private Function pGenerarConsulta(mIdParte As String, mIdReceta As String, fTipo As e_PROGRAMA, dFecha As Date, Optional fAddRegistro As Boolean = False, Optional fAddRegistroSinPrograma As Boolean = False) As String    '--
    Dim nSQL As String
    
    If fTipo = E_INSUMO Then
        If fAddRegistro = True Then '--NUEVO
            If fAddRegistroSinPrograma = False Then
                nSQL = "SELECT " + mIdParte + " AS idparte, pro_programadet.idrec, pro_recetains.iditem, pro_recetains.idunimed, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro AS unid, pro_recetains!canpro*pro_programadet!canpro AS canprog, '' AS canteo, '' AS canreal, '' AS dif " _
                    + vbCr + " FROM (pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades RIGHT JOIN ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON mae_unidades.id = pro_recetains.idunimed) ON pro_receta.id = pro_recetains.idrec " _
                    + vbCr + " WHERE (((pro_programadet.idrec) = " + mIdReceta + ") And ((pro_programadet.dia) = CDATE('" + CStr(dFecha) + "') )) " _
                    + vbCr + " ORDER BY mae_tipoproducto.descripcion ASC, alm_inventario.descripcion;"
            Else '--DIRECTAMENTE DE RECETAS(SIN PROGRAMA)
                nSQL = "SELECT " + mIdParte + " AS idparte,pro_receta.id AS idrec,pro_recetains.iditem, pro_recetains.idunimed, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro AS unid, 0 AS canprog, '' AS canteo, '' AS canreal, '' AS dif " _
                    + vbCr + " FROM pro_receta INNER JOIN (mae_unidades RIGHT JOIN ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON mae_unidades.id = pro_recetains.idunimed) ON pro_receta.id = pro_recetains.idrec " _
                    + vbCr + " WHERE (((pro_receta.ID) = " + mIdReceta + ")) " _
                    + vbCr + " ORDER BY mae_tipoproducto.descripcion, alm_inventario.descripcion;"
            End If
        Else '--CONSULTA O MODIFICAR
            nSQL = "SELECT pro_producciondet.numparte AS idparte, pro_producciondetins.idrec, pro_producciondetins.iditem, pro_producciondetins.idunimed, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_producciondetins.canpro AS unid, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_producciondetins.canpro) AS canprog, IIf(pro_producciondet.cantidad Is Null,0,(pro_producciondet.cantidad*pro_producciondetins.canpro)) AS canteo, pro_producciondetins.canutil AS canreal, [canteo]-pro_producciondetins!canutil AS dif " _
                + vbCr + " FROM (pro_producciondet LEFT JOIN pro_programadet ON (pro_producciondet.idpro = pro_programadet.idpro) AND (pro_producciondet.idrec = pro_programadet.idrec)) INNER JOIN (mae_tipoproducto RIGHT JOIN (((pro_producciondetins LEFT JOIN mae_unidades ON pro_producciondetins.idunimed = mae_unidades.id) LEFT JOIN pro_recetains ON (pro_producciondetins.idrec = pro_recetains.idrec) AND (pro_producciondetins.iditem = pro_recetains.iditem)) LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) ON mae_tipoproducto.id = alm_inventario.tippro) ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
                + vbCr + " GROUP BY pro_producciondet.numparte, pro_producciondetins.idrec, pro_producciondetins.iditem, pro_producciondetins.idunimed, mae_tipoproducto.descripcion, alm_inventario.descripcion, mae_unidades.abrev, pro_producciondetins.canpro, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_producciondetins.canpro), IIf(pro_producciondet.cantidad Is Null,0,(pro_producciondet.cantidad*pro_producciondetins.canpro)), pro_producciondetins.canutil " _
                + vbCr + " HAVING (((pro_producciondet.numparte) = '" + mIdParte + "') And ((pro_producciondetins.idrec) = " + mIdReceta + ")) " _
                + vbCr + " ORDER BY mae_tipoproducto.descripcion, alm_inventario.descripcion;"
        End If
    ElseIf fTipo = e_TAREA Then
        If fAddRegistro = True Then
            If fAddRegistroSinPrograma = False Then
                nSQL = "SELECT " + mIdParte + " AS idparte, pro_programadet.idrec,pro_recetatar.orden as corr, pro_recetatar.idtar, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad AS unid, pro_programadet!canpro*pro_recetatar!cantidad AS canprog,  '' AS horini, '' AS horfin, 0 as canper " _
                    + vbCr + " FROM  pro_tareas INNER JOIN ((pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades INNER JOIN pro_recetatar ON mae_unidades.id = pro_recetatar.idunimed) ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar " _
                    + vbCr + " WHERE (((pro_programadet.idrec) = " + mIdReceta + ") And ((pro_programadet.dia) = CDATE('" + CStr(dFecha) + "') )) " _
                    + vbCr + " ORDER BY pro_recetatar.orden;"
            Else '--DIRECTAMENTE DE RECETAS
                nSQL = "SELECT " + mIdParte + " AS idparte,pro_receta.id AS idrec, pro_recetatar.idtar, pro_recetatar.orden as corr, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad AS unid, 0 AS canprog,'' AS horini, '' AS horfin, 0 AS canper " _
                    + vbCr + " FROM pro_tareas INNER JOIN (pro_receta INNER JOIN (mae_unidades INNER JOIN pro_recetatar ON mae_unidades.id = pro_recetatar.idunimed) ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar " _
                    + vbCr + " WHERE (((pro_receta.ID) = " + mIdReceta + ")) " _
                    + vbCr + " ORDER BY pro_recetatar.orden;"
            End If
        Else
           nSQL = "SELECT pro_producciondet.numparte as idparte ,pro_producciondettar.idrec, pro_recetatar.idtar,pro_producciondettar.corr, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad AS unid, IIF(pro_programadet.canpro IS NULL,0,pro_programadet.canpro*pro_recetatar.cantidad) AS canprog, Format([pro_producciondettar].[horini],'hh:nn AM/PM') AS horini, Format([pro_producciondettar].[horfin],'hh:nn AM/PM') AS horfin, pro_producciondettar.canper " _
                + vbCr + " FROM (pro_producciondet LEFT JOIN pro_programadet ON (pro_producciondet.idpro = pro_programadet.idpro) AND (pro_producciondet.idrec = pro_programadet.idrec)) INNER JOIN (pro_tareas RIGHT JOIN ((pro_producciondettar INNER JOIN pro_recetatar ON (pro_producciondettar.idtar = pro_recetatar.idtar) AND (pro_producciondettar.idrec = pro_recetatar.idrec)) LEFT JOIN mae_unidades ON pro_producciondettar.idunimed = mae_unidades.id) ON pro_tareas.id = pro_producciondettar.idtar) ON (pro_producciondet.idpro = pro_producciondettar.idpro) AND (pro_producciondet.numparte = pro_producciondettar.numparte) AND (pro_producciondet.idrec = pro_producciondettar.idrec) " _
                + vbCr + " GROUP BY pro_producciondet.numparte, pro_producciondettar.idrec, pro_recetatar.idtar,pro_producciondettar.corr, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_recetatar.cantidad), pro_producciondettar.canper, pro_producciondettar.horini, pro_producciondettar.horfin, pro_recetatar.orden, pro_programadet.dia, pro_recetatar.orden " _
                + vbCr + " HAVING (((pro_producciondet.numparte) = '" + mIdParte + "') And ((pro_producciondettar.idrec) = " + mIdReceta + ")) " _
                + vbCr + " ORDER BY pro_tareas.descripcion, pro_recetatar.orden, pro_recetatar.orden;"
        End If
    ElseIf fTipo = e_EQUIPO Then
        '--FALTA DEFINIR
        '-------------
        '----------------
    End If
    pGenerarConsulta = nSQL
End Function

'*****************************************************************************************************
'* Nombre           : DATOS_TMP_ADD
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESTA FUNCION CARGA LOS DATOS RELACIONADO A LOS INSUMOS, TAREAS, PARA LUEGO SER
'*                    MOSTRADO EN EL GRID DE INSUMOS Y TAREAS
'* Parametros       : NOMBRE                   |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    mIdParte                 |  String      |
'*                    mIdReceta                |  String      |
'*                    fTipo                    |  e_PROGRAMA  |
'*                    fAddRegistro             |  Boolean     |
'*                    fAddRegistroSinPrograma  |  Boolean     |
'* Devuelve         :
'*****************************************************************************************************
Private Sub DATOS_TMP_ADD(mIdParte As String, mIdReceta As String, fTipo As e_PROGRAMA, _
                            Optional fAddRegistro As Boolean = False, _
                            Optional fAddRegistroSinPrograma As Boolean = False)
                            
    On Error GoTo error
    Dim RST_ORIGEN As New ADODB.Recordset
    Dim nSQL As String
    
    Me.MousePointer = vbHourglass
    nSQL = pGenerarConsulta(mIdParte, mIdReceta, fTipo, TxtFecha(0).valor, fAddRegistro, fAddRegistroSinPrograma)
    If nSQL = "" Then GoTo SALIR
    RST_Busq RST_ORIGEN, nSQL, xCon
    If RST_ORIGEN.State = 0 Then GoTo SALIR
    
    If RST_ORIGEN.EOF = True Or RST_ORIGEN.BOF = True Or RST_ORIGEN.RecordCount = 0 Then GoTo SALIR:
    
    If fTipo = E_INSUMO Then
        CARGAR_RST_TMP RST_INSUMO, RST_ORIGEN
    ElseIf fTipo = e_TAREA Then
        CARGAR_RST_TMP RST_TAREA, RST_ORIGEN
    ElseIf fTipo = e_EQUIPO Then
    End If

SALIR:
    Me.MousePointer = vbDefault
    Set RST_ORIGEN = Nothing
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Set RST_ORIGEN = Nothing
    SHOW_ERROR Me.Name, "DATOS_TMP_ADD"
End Sub

'*****************************************************************************************************
'* Nombre           : DATOS_TMP_DEL
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINAR DATOS DEL TEMPORAL
'* Parametros       : NOMBRE      |  TIPO            |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    mIdParte    |  String          |
'*                    mIdReceta   |  String          |
'*                    RST_TMP     |  ADODB.Recordset |
'* Devuelve         :
'*****************************************************************************************************
Private Sub DATOS_TMP_DEL(mIdParte As String, mIdReceta As String, RST_TMP As ADODB.Recordset)
    If mIdReceta = "" Then Exit Sub
    
    RST_TMP.Filter = "idparte= " + mIdParte + " AND idrec=" + mIdReceta
    
    If RST_TMP.RecordCount = 0 Then Exit Sub
    
    RST_TMP.MoveFirst
    
    Do While Not RST_TMP.EOF
        RST_TMP.Delete
        RST_TMP.MoveNext
    Loop
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel()
    Dim IDPROCORR_ As Double
    Dim xRs As New ADODB.Recordset
    Dim MENSAJE_ As String
    
    If Fg1.Row < 0 Then Exit Sub
    If Fg1.Row = 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation
        Exit Sub
    End If
    
    '*****************************************************************
    If NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then
        MsgBox "El registro no se puede eliminar debido a su estado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
'    If Not verificarCambioEstado(Fg1.TextMatrix(Fg1.Row, COLUMNACORR_), MENSAJE_) Then
'        MsgBox MENSAJE_, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
    '*****************************************************************
    
    If MsgBox("Seguro desea eliminar el Producto", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    If Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_) <> "" Then
        ' ELIMINANDO LOS REGISTROS DE RECORSET TEMPORAL
        DATOS_TMP_DEL Fg1.TextMatrix(Fg1.Row, COLUMNAIDPARTE_), Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_), RST_INSUMO
        DATOS_TMP_DEL Fg1.TextMatrix(Fg1.Row, COLUMNAIDPARTE_), Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_), RST_TAREA
    End If
    
    Label2.Caption = ""
    
    ' SE ELIMINAN LOS REGISTROS RELACIONADOS
    IDPROCORR_ = NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNACORR_))
    ' SOLICITUD DE MATERIALES
'    cSQL = "SELECT pro_ordenproddet.idprocorr, pro_ordenproddet.id " _
'            + vbCr + "FROM pro_ordenproddet " _
'            + vbCr + "WHERE (((pro_ordenproddet.idprocorr)=" & IDPROCORR_ & "));"
'
'    Set xRs = Nothing
'    RST_Busq xRs, cSQL, xCon
'
'    If xRs.State = 0 Then GoTo SIGUIENTE_
'    If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
'
'    xRs.MoveFirst
'    While Not xRs.EOF
'        xCon.Execute "DELETE * FROM pro_ordenproddetins WHERE idorddet = " & NulosN(xRs("id"))
'        xCon.Execute "DELETE * FROM pro_ordenproddet WHERE id = " & NulosN(xRs("id"))
'        xRs.MoveNext
'    Wend
    CAMBIOGRABAR_ = -1
    
    ' REGISTRO DE PLANILLAS
    cSQL = "SELECT pro_controltardet.idprocorr, pro_controltardet.idctr, pro_controltardet.corr " _
            + vbCr + "FROM pro_controltardet " _
            + vbCr + "WHERE (((pro_controltardet.idprocorr)=" & IDPROCORR_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then GoTo SIGUIENTE_
    If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
    
    xRs.MoveFirst
    While Not xRs.EOF
        xCon.Execute "DELETE * FROM pro_controltardettar WHERE idctr = " & NulosN(xRs("idctr")) & " And corr=" & NulosN(xRs("corr"))
        xCon.Execute "DELETE * FROM pro_controltardetgr WHERE idctr = " & NulosN(xRs("idctr")) & " And corr=" & NulosN(xRs("corr"))
        xCon.Execute "DELETE * FROM pro_controltardet WHERE idctr = " & NulosN(xRs("idctr")) & " And corr=" & NulosN(xRs("corr"))
        xRs.MoveNext
    Wend
    CAMBIOGRABAR_ = -1
    
    ' ENTRADAS Y SALIDAS
    cSQL = "SELECT alm_ingreso.idprocorr, alm_ingreso.idorddet, alm_ingreso.id " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idprocorr)=" & IDPROCORR_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then GoTo SIGUIENTE_
    If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
    
    xRs.MoveFirst
    While Not xRs.EOF
        xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & NulosN(xRs("id"))
        xCon.Execute "DELETE * FROM alm_ingreso WHERE id = " & NulosN(xRs("id"))
        xRs.MoveNext
    Wend
    CAMBIOGRABAR_ = -1
SIGUIENTE_:
        
    If NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNANUMPROD_)) = NUMEROPROD_ - 1 Then NUMEROPROD_ = NUMEROPROD_ - 1
    
    ' ELIMINAR EL PRODUCTO
    Fg1.RemoveItem (Fg1.Row)
    If Fg1.Rows > 1 Then Fg1.Row = 1
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 ' AGREGAR PROGRAMA
            pRegistroAdd False, True
        
        Case 1 ' ELIMINAR REGISTROS AGREGADOS
            pRegistroDel
        
        Case 2
            pRegistroAdd True, True
        
        Case 3
            pRegistroAdd True, False
    End Select
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            If Index = 0 Then PopupMenu menu3
            If Index = 1 Then PopupMenu menu2
        End If
    End If
End Sub

Private Sub Dg_AfterColEdit(Index As Integer, ByVal ColIndex As Integer)
    On Error GoTo error
    If Index = 0 Then
        If ColIndex = Dg(Index).Columns("canreal").ColIndex Then
            If Dg(Index).Columns("canreal") = "" Then Exit Sub
            
            If IsNumeric(Dg(Index).Columns("canreal")) = False Then
                MsgBox "La Cantidad ingresada no es numérico", vbExclamation, xTitulo
                If Index = 0 Then
                    RST_INSUMO.Fields("canreal") = ""
                    RST_INSUMO.UpdateBatch
                ElseIf Index = 1 Then
                    RST_TAREA.Fields("canreal") = ""
                    RST_TAREA.UpdateBatch
                End If
                Exit Sub
            End If
            
            ' CALCULANDO LA DIFERENMCIA
            If Index = 0 Then
                If NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNATOTPROD_)) = 0 Then
                    RST_INSUMO.Fields("dif") = 0
                Else
                    RST_INSUMO.Fields("dif") = NulosN(RST_INSUMO.Fields("canteo")) - NulosN(RST_INSUMO.Fields("canreal"))
                End If
                RST_INSUMO.UpdateBatch
            ElseIf Index = 2 Then
                
            End If
        End If
    ElseIf Index = 1 Then
        If Trim(Dg(1).Columns(ColIndex)) = "" Then Exit Sub
        
        If ColIndex = Dg(1).Columns("horini").ColIndex Or ColIndex = Dg(1).Columns("horfin").ColIndex Or ColIndex = Dg(1).Columns("canper").ColIndex Then
            If ColIndex <> Dg(1).Columns("canper").ColIndex Then '--es una hora
                ' VALIDAR QUE LA HORA DE INICIO Y FIN DE LA PRODUCCION ESTEN INGRESADOS, SI NO ES ASI =>> SALIR
                If IsDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_)) = False Or IsDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORFIN_)) = False Then
                    MsgBox "Falta ingresar la hora " + IIf(IsDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_)) = False, "Inicial", "Final") + " de la producción" + vbCr + _
                    "Ingrese la información solicitada para continuar", vbExclamation, xTitulo
                    Dg(1).Columns(ColIndex) = "":  Fg1.SetFocus:   Exit Sub
                End If
                
                ' VALIDAR QUE EXISTA EL SEPARADOR DE HORA Y MINUTO (:)
                If InStr(Dg(1).Columns(ColIndex), ":") = 0 Then Dg(1).Columns(ColIndex) = Mid(Dg(1).Columns(ColIndex), 1, 2) + ":" + Mid(Dg(1).Columns(ColIndex), 3)
                
                ' VER SI ES HORA CORRECTA
                If IsDate(Dg(1).Columns(ColIndex)) = False Then
                    MsgBox "El valor ingresado no es una Hora correcta", vbCritical, xTitulo
                    Dg(1).Columns(ColIndex) = ""
                Else
                    ' VALIDAR QUE LA HORA ESTE EN EL INTERVALO DE LA HORA INICIAL Y FINAL DE LA PRODUCCION
                    'HORA < HORA_INICIO_PROD Ó HORA > HORA_FINAL_PROD =>> MONTRAR MSG
                If (CDate(Dg(1).Columns(ColIndex)) < CDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_))) Or (CDate(Dg(1).Columns(ColIndex)) > CDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORFIN_))) Then
                    MsgBox "La hora ingresada es " + IIf(CDate(Dg(1).Columns(ColIndex)) < CDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_)), " inferior a la hora inicial", "superior a la hora final") + " de la Producción" + vbCr + _
                            IIf(CDate(Dg(1).Columns(ColIndex)) < CDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_)), "Inicio de Producción: ", "Término de Producción: ") + " " + IIf(CDate(Dg(1).Columns(ColIndex)) < CDate(Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_)), Fg1.TextMatrix(Fg1.Row, COLUMNAHORINI_), Fg1.TextMatrix(Fg1.Row, COLUMNAHORFIN_)) + vbCr + _
                            "Hora Ingresada:          " + Format(Dg(1).Columns(ColIndex), FORMAT_HORA_SIN_SEGUNDO) + vbCr + _
                            "Modifique la hora ingresada", vbExclamation, xTitulo
                    Dg(1).Columns(ColIndex) = "":     Exit Sub
                End If
                    If IsDate(RST_TAREA.Fields("horini")) = True And IsDate(RST_TAREA.Fields("horfin")) = True Then
                        ' VALIDAR SI LA HORA INICIAL ES MAYOR O IGUAL AL FINAL =>> SALIR
                        If CDate(RST_TAREA.Fields("horini")) >= CDate(RST_TAREA.Fields("horfin")) Then
                            MsgBox "La hora " + IIf(ColIndex = 5, "Inicial debe ser menor ", "Final debe ser mayor") + " a la hora " + IIf(ColIndex = 5, "Final", "Inicial"), vbExclamation, xTitulo
                            Dg(1).Columns(ColIndex) = "":      Exit Sub
                       End If
                    End If
                    ' DAR FORMATO A LA HORA INGRESADA
                    Dg(1).Columns(ColIndex) = Format(Dg(1).Columns(ColIndex), FORMAT_HORA_SIN_SEGUNDO)
                End If
            Else
                If IsNumeric(RST_TAREA.Fields("canper")) = False Then
                    MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                    RST_TAREA.Fields("canper") = ""
                Else
                    RST_TAREA.Fields("canper") = CInt(RST_TAREA.Fields("canper"))
                End If
            End If
        End If
    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Dg_AfterColEdit(" + CStr(Index) + ")"
End Sub

Private Sub Dg_BeforeColEdit(Index As Integer, ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then Cancel = 1
End Sub

Private Sub Dg_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
    If Index <> 0 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    '************************************************************************
    ' Se verifica el estado
    If Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_) >= ESTADOAPROBADO_ Then Exit Sub
    '************************************************************************
    
    If Dg(Index).Columns("descripcion").ColIndex <> ColIndex Then Exit Sub

    ' GENERAR EL WHERE DE LOS ID'S DE CUENTA PARA QUE NO SE REPITAN
    Dim SQL_ID As String
    Dim nSQL As String
    Dim V_POSICION As Variant
    
    If RST_INSUMO.EOF = False Or RST_INSUMO.BOF = False Or RST_INSUMO.RecordCount <> 0 Then
        V_POSICION = RST_INSUMO.Bookmark
        Dg(Index).MoveFirst
    End If
    
    Do While Not RST_INSUMO.EOF
        If RST_INSUMO.Fields("iditem") & "" <> "" Then
            SQL_ID = SQL_ID + CStr(RST_INSUMO.Fields("iditem") & "") + ","
        End If
        RST_INSUMO.MoveNext
    Loop
    
    If SQL_ID <> "" Then SQL_ID = " AND alm_inventario.id NOT IN (" + Left(SQL_ID, Len(SQL_ID) - 1) + ") "
    If RST_INSUMO.RecordCount <> 0 Then RST_INSUMO.Bookmark = V_POSICION
    
    Dim xRs As New ADODB.Recordset
    ReDim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000": xCampos(0, 3) = "C":
    xCampos(1, 0) = "U.M.":           xCampos(1, 1) = "abrev":        xCampos(1, 2) = "500":  xCampos(1, 3) = "C":
    xCampos(2, 0) = "Tipo Producto":  xCampos(2, 1) = "tipprodesc":   xCampos(2, 2) = "1500": xCampos(2, 3) = "C":

    nSQL = "SELECT alm_inventario.id, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS tipprodesc " _
        + vbCr + " FROM mae_unidades INNER JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed " _
        + vbCr + " WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro) In (1,3,4))) " + SQL_ID _
        + vbCr + " ORDER BY alm_inventario.descripcion;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Items", "descripcion", "descripcion", Principio, ""

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    
    RST_INSUMO.Fields("idparte") = Fg1.TextMatrix(Fg1.Row, COLUMNAIDPARTE_)
    RST_INSUMO.Fields("idrec") = IIf(Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_) = "", "-999", Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_))
    RST_INSUMO.Fields("iditem") = xRs.Fields("id") & ""
    RST_INSUMO.Fields("descripcion") = xRs.Fields("descripcion") & ""
    RST_INSUMO.Fields("tipprodesc") = xRs.Fields("tipprodesc") & ""
    RST_INSUMO.Fields("abrev") = xRs.Fields("abrev") & ""
    RST_INSUMO.Fields("idunimed") = xRs.Fields("idunimed") & ""
    RST_INSUMO.Fields("unid") = 0
    RST_INSUMO.Fields("canprog") = 0
    RST_INSUMO.Fields("canteo") = 0
    RST_INSUMO.Update
    Agregando = False
    Set xRs = Nothing
    Exit Sub

SALIR:
    Set xRs = Nothing
    Agregando = False

error:
    SHOW_ERROR Me.Name, "Dg_ButtonClick(" + CStr(Index) + ")"
End Sub

Private Sub Dg_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
End Sub

Private Sub Dg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    ' INDEX <> INSUMO SALIR
    If KeyCode = 117 Then
        ' F6 ENFOCAR EN EL GRID DE PRODUCTOS
        Fg1.SetFocus
    End If
    
    If Index <> 0 Then Exit Sub
    
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then              ' F3 = Agregar Item
        If RST_INSUMO.EOF = False Or RST_INSUMO.BOF = False Or RST_INSUMO.RecordCount <> 0 Then
            Dg(Index).MoveLast
            If NulosN(RST_INSUMO.Fields("iditem")) = 0 Then Exit Sub
        End If

        RST_INSUMO.AddNew
        RST_INSUMO.Fields("idparte") = Fg1.TextMatrix(Fg1.Row, COLUMNAIDPARTE_)
        RST_INSUMO.Fields("idrec") = IIf(Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_) = "", "-999", Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_))
        
        RST_INSUMO.MoveLast
        ' CARGAR EL FORMULARIO DE SELECCION DE ITEM PARA AGREGAR UN REGISTRO AL GRID
        Dg_ButtonClick 0, 4
    End If
    
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        If RST_INSUMO.RecordCount = 0 Then Exit Sub
        If MsgBox("Seguro desea Eliminar el insumo seleccionado", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        If RST_INSUMO.BOF = False Or RST_INSUMO.EOF = False Or RST_INSUMO.RecordCount <> 0 Then RST_INSUMO.Delete
        ' INICIALIZANDO OTRA VEZ EL FILTRO SEGUN PRODUCTO
        Fg1_RowColChange
    End If
End Sub

Private Sub Dg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> 0 Then Exit Sub
    
    If Button = 2 Then
        If QueHace <> 3 Then
            ' Se verifica si esta o no Aprobada
            If Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_) >= ESTADOAPROBADO_ Then Exit Sub
            PopupMenu Menu5
        End If
    End If
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_FilterChange()
    TDB_FiltroGenerar Dg3, RstFrm
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If NulosN(Fg1.TextMatrix(Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then
        Cancel = True
    End If
    
    Select Case Col
        Case COLUMNAESTADO_
            ' Se llena el estado anterior
            ESTADOANTERIOR_ = NulosN(Fg1.TextMatrix(Row, Col))
            
    End Select
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim RST_TMP As New ADODB.Recordset
    
    On Error GoTo error
    
    If Agregando = True Then Exit Sub
    If Row = 0 Then Exit Sub
    
    If NulosN(Fg1.TextMatrix(Row, COLUMNAIDREC_)) = 0 Then
        MsgBox "Ingrese La Receta del Producto:" + vbCr + _
        "Producto:        " + Fg1.TextMatrix(Row, COLUMNAPRODUCTO_) & "", vbExclamation, xTitulo
        Fg1.TextMatrix(Row, Col) = ""
        Fg1.Col = COLUMNARECETA_: Exit Sub
    End If
    
    Select Case Col
        Case COLUMNANUMPROD_
            If Fg1.TextMatrix(Row, COLUMNAIDREC_) = "" Then Exit Sub
            RST_Busq RST_TMP, "select numparte from pro_producciondet where numparte='" + Format(NulosN(Fg1.TextMatrix(Row, COLUMNANUMPROD_)), FORMAT_NUM_PARTE) + "' ;", xCon
            If RST_TMP.EOF = False And RST_TMP.BOF = False And RST_TMP.RecordCount <> 0 Then
                Set RST_TMP = Nothing
                If MsgBox("El número de producción ya existe, desea continuar?", vbQuestion + vbYesNo, xTitulo) = vbNo Then
                    Fg1.TextMatrix(Row, Col) = ""
                    Fg1.Col = COLUMNANUMPROD_:    Exit Sub
                End If
            End If
            Set RST_TMP = Nothing
        
        Case COLUMNATOTPROD_
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_CANTIDAD)
                ' ACTUALIZAR LA CANTIDAD TEORICA
                If RST_INSUMO.EOF = False Or RST_INSUMO.BOF = False Or RST_INSUMO.RecordCount <> 0 Then RST_INSUMO.MoveFirst
                Do While Not RST_INSUMO.EOF
                    RST_INSUMO.Fields("canteo") = NulosN(RST_INSUMO.Fields("unid")) * NulosN(Fg1.TextMatrix(Row, Col))
                    RST_INSUMO.MoveNext
                Loop
            End If
            
        Case COLUMNAHORINI_, COLUMNAHORFIN_
            If IsDate(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es una Hora correcta", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
'                If IsDate(Fg1.TextMatrix(Row, COLUMNAHORINI_)) = True And IsDate(Fg1.TextMatrix(Row, COLUMNAHORFIN_)) = True Then '--HORA INICIO
'                    If CDate(Fg1.TextMatrix(Row, COLUMNAHORINI_)) >= CDate(Fg1.TextMatrix(Row, COLUMNAHORFIN_)) Then
'                        MsgBox "La hora " + IIf(Col = COLUMNAHORINI_, "Inicial debe ser menor ", "Final debe ser mayor") + " a la hora " + IIf(Col = COLUMNAHORINI_, "Final", "Inicial"), vbExclamation, xTitulo
'                        Fg1.TextMatrix(Row, Col) = "":  Exit Sub
'                   End If
'                End If
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            End If
    End Select
    Exit Sub

error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "Fg1_CellChanged"
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> COLUMNARECETA_ And Col <> COLUMNARESPONSABLE_ _
            And Col <> COLUMNATURNO_ And Col <> COLUMNAORDENPROD_ Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String
    Dim nOrden As String
    
    If NulosN(Fg1.TextMatrix(Row, COLUMNAIDREC_)) = 0 And Col <> COLUMNARECETA_ Then
        MsgBox "Ingrese La Receta del Producto:" + vbCr + _
        "Producto:        " + Fg1.TextMatrix(Row, COLUMNAPRODUCTO_) & "", vbExclamation, xTitulo
        Fg1.TextMatrix(Row, Col) = ""
        Fg1.Col = COLUMNARECETA_: Exit Sub
    End If
    
    Select Case Col
        Case COLUMNARECETA_  ' DE LAS RECETAS
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Descripción":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Código":          xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "U.M.":            xCampos(2, 1) = "abrev":         xCampos(2, 2) = "500":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
            nSQL = "SELECT pro_receta.id AS idrec, pro_receta.descripcion as nombre, pro_receta.codrec, mae_unidades.abrev, pro_receta.idunimed " _
                    + vbCr + " FROM alm_inventario INNER JOIN (pro_receta INNER JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem " _
                    + vbCr + " WHERE (((alm_inventario.id) = " + CStr(Fg1.TextMatrix(Row, COLUMNAIDITEM_)) + ")) " _
                    + vbCr + " ORDER BY pro_receta.descripcion;"
            nTitulo = "Buscando Recetas"
            nOrden = "nombre"
        
        Case COLUMNARESPONSABLE_  ' DEL RESPONSABLE
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Nombre":       xCampos(0, 1) = "nombre":       xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "DNI":          xCampos(1, 1) = "numdoc":       xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
            nSQL = "SELECT pro_emp.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.numdoc " _
                + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
                + vbCr + " Where (((pro_empdet.idfun) = 3)) " _
                + vbCr + " ORDER BY pla_empleados.apepat;"
            nTitulo = "Buscando Responsables de Producción"
            nOrden = "nombre"
    
        Case COLUMNATURNO_ ' DEL TURNO
            ReDim xCampos(1, 4) As String
            xCampos(0, 0) = "Turno":       xCampos(0, 1) = "nombre":       xCampos(0, 2) = "3500":    xCampos(0, 3) = "C"
            nSQL = "SELECT mae_turnos.id, mae_turnos.descripcion as nombre FROM mae_turnos;"
            nTitulo = "Buscando Turnos"
            nOrden = "nombre"
        '*****************************************************************************************************
        Case COLUMNAORDENPROD_  ' ORDEN DE PRODUCCIÓN
            ReDim xCampos(6, 4) As String
            xCampos(0, 0) = "Fecha.":           xCampos(0, 1) = "fchpro":           xCampos(0, 2) = "1000":          xCampos(0, 3) = "C"
            xCampos(1, 0) = "Num. Ord.":        xCampos(1, 1) = "numdoc":           xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Ítem":             xCampos(2, 1) = "item":             xCampos(2, 2) = "1900":         xCampos(2, 3) = "C"
            xCampos(3, 0) = "Responsable":      xCampos(3, 1) = "desresp":          xCampos(3, 2) = "1900":         xCampos(3, 3) = "C"
            xCampos(4, 0) = "Cantidad":         xCampos(4, 1) = "cantidad":         xCampos(4, 2) = "900":          xCampos(4, 3) = "N"
            xCampos(5, 0) = "Hor.Ini.":         xCampos(5, 1) = "horini":           xCampos(5, 2) = "1100":          xCampos(5, 3) = "C"

            nSQL = "SELECT pro_ordenprod.id, pro_ordenprod.fchpro, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numdoc, alm_inventario.descripcion AS item, pla_empleados.nombre AS desresp, pro_ordenprod.cantidad, pro_ordenprod.horini, pro_ordenprod.horfin " _
                + vbCr + "FROM ((pro_ordenprod LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN pla_empleados ON pro_ordenprod.idresp = pla_empleados.id " _
                + vbCr + "WHERE (((pro_ordenprod.idrec)=" & NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_)) & ") AND ((pro_ordenprod.estado)=" & ESTADOPROCESADO_ & ") AND ((pro_ordenprod.idmes) In (" & mMesActivo & "," & mMesActivo - 1 & ")));"
            
'            nSQL = "SELECT pro_cronogramadet.numprod, pro_cronogramadet.id, alm_inventario.descripcion AS nombre, pro_cronogramadet.fchpro, pro_cronogramadet.horpro, pro_cronogramadet.cantidad, mae_unidades.abrev, pro_cronogramadet.iditem, pro_receta.id AS idrec, pro_receta.prirec, pro_receta.codrec, alm_inventario.idunimed, pla_empleados.nombre AS nomsup " _
'                    + vbCr + "FROM ((((pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) LEFT JOIN pro_cronogramatarea ON pro_cronogramadet.id = pro_cronogramatarea.idcrdet) LEFT JOIN pla_empleados ON pro_cronogramatarea.idresp = pla_empleados.id " _
'                    + vbCr + "GROUP BY pro_cronogramadet.numprod, pro_cronogramadet.id, alm_inventario.descripcion, pro_cronogramadet.fchpro, pro_cronogramadet.horpro, pro_cronogramadet.cantidad, mae_unidades.abrev, pro_cronogramadet.iditem, pro_receta.id, pro_receta.prirec, pro_receta.codrec, alm_inventario.idunimed, pla_empleados.nombre " _
'                    + vbCr + "HAVING (((pro_cronogramadet.numprod)<>'') AND ((pro_receta.id)=" & NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_)) & ") AND ((pro_receta.prirec)=1));"
                    
'            nSQL = "SELECT pro_cronogramadet.numprod, pro_cronogramadet.id, alm_inventario.descripcion As nombre, pro_cronogramadet.fchpro, pro_cronogramadet.horpro, pro_cronogramadet.cantidad, mae_unidades.abrev, pro_cronogramadet.iditem, pro_receta.id AS idrec, pro_receta.prirec, pro_receta.codrec, alm_inventario.idunimed " _
'                    + vbCr + "FROM ((pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem " _
'                    + vbCr + "WHERE (((pro_cronogramadet.numprod)<>'') AND ((pro_receta.prirec)=1));"
                
            nTitulo = "Buscando Programación"
            nOrden = "fchpro"
        '*****************************************************************************************************
    End Select

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, "nombre", Principio, ""

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    
    If Col = COLUMNARECETA_ Then ' DE LA RECETA
        If GRID_BUSCAR_VALOR(Fg1, COLUMNARECETA_, xRs.Fields("codrec") & "", False, , Row) <> "-1" Then
            MsgBox "La receta ya existe" + vbCr + "Seleccione otra Receta o Elimine el Producto", vbExclamation, xTitulo
            GoTo SALIR:
        Else
            ' ELIMINAR RECETA ANTERIOR
            DATOS_TMP_DEL Fg1.TextMatrix(Row, COLUMNAIDPARTE_), Fg1.TextMatrix(Row, COLUMNAIDREC_), RST_INSUMO
            DATOS_TMP_DEL Fg1.TextMatrix(Row, COLUMNAIDPARTE_), Fg1.TextMatrix(Row, COLUMNAIDREC_), RST_TAREA
            
            Fg1.TextMatrix(Row, COLUMNARECETA_) = xRs.Fields("codrec") & ""
            Fg1.TextMatrix(Row, COLUMNAUM_) = xRs.Fields("abrev") & ""
            Fg1.TextMatrix(Row, COLUMNAIDREC_) = xRs.Fields("idrec") & ""
            Fg1.TextMatrix(Row, COLUMNAIDUNID_) = xRs.Fields("idunimed") & ""
            
            ' AGREGANDO NUEVA RECETA
            DATOS_TMP_ADD Fg1.TextMatrix(Row, COLUMNAIDPARTE_), xRs.Fields("idrec"), E_INSUMO, True, True
            DATOS_TMP_ADD Fg1.TextMatrix(Row, COLUMNAIDPARTE_), xRs.Fields("idrec"), e_TAREA, True, True
            
            If RST_INSUMO.EOF = False Or RST_INSUMO.BOF = False Or RST_INSUMO.RecordCount <> 0 Then RST_INSUMO.MoveFirst
            Do While Not RST_INSUMO.EOF
                RST_INSUMO.Fields("canteo") = NulosN(RST_INSUMO.Fields("unid")) * NulosN(Fg1.TextMatrix(Row, COLUMNATOTPROD_))
                RST_INSUMO.MoveNext
            Loop
        End If
        Fg1.Col = COLUMNATOTPROD_
    ElseIf Col = COLUMNARESPONSABLE_ Then ' DEL RESPONSABLE DE PRODUCCION
        Fg1.TextMatrix(Row, COLUMNARESPONSABLE_) = NulosC(xRs.Fields("nombre"))
        Fg1.TextMatrix(Row, COLUMNAIDRES_) = NulosN(xRs.Fields("id"))
        Fg1.Col = COLUMNATURNO_
    ElseIf Col = COLUMNATURNO_ Then ' DE LA UNIDAD DE MEDIDA
        Fg1.TextMatrix(Row, COLUMNATURNO_) = NulosC(xRs.Fields("nombre"))
        Fg1.TextMatrix(Row, COLUMNAIDTURNO_) = NulosN(xRs.Fields("id"))
    '***********************************************************************************************
    ElseIf Col = COLUMNAORDENPROD_ Then ' NUMERO DE PRODUCCION
        Fg1.TextMatrix(Row, COLUMNAORDENPROD_) = NulosC(xRs.Fields("numdoc"))
        Fg1.TextMatrix(Row, COLUMNAIDORD_) = NulosN(xRs.Fields("id"))
    '***********************************************************************************************
    End If
    Fg1.SetFocus
    Agregando = False
    Set xRs = Nothing
    Exit Sub
    
SALIR:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Function cambiarEstadoRelacionados(IDREGPROD_ As Double, ESTADO_ As Double) As Boolean
    On Error GoTo ERROR_
    Dim ID_ As Double
    
    ' Solicitud de Materiales
    cSQL = "UPDATE pro_ordenproddet SET pro_ordenproddet.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((pro_ordenproddet.idprocorr)=" & IDREGPROD_ & "));"
    
    xCon.Execute cSQL
    
    ' Registros de Planillas
    cSQL = "UPDATE pro_controltardet SET pro_controltardet.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((pro_controltardet.idprocorr)=" & IDREGPROD_ & "));"
    
    xCon.Execute cSQL
    
    ' Salidas de Almacen
    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((alm_ingreso.idprocorr)=" & IDREGPROD_ & "));"

    xCon.Execute cSQL
    
    ' GRABAMOS LOS MOVIMIENTOS
    ' SOLICITUD DE MATERIALES
    ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "idord", "pro_ordenproddet", "N", xCon)
    GrabarOperacion xIdUsuario, 54, 7, xHorIni, Time, Date, xCon, ID_
    ' REGISTRO DE PLANILLA
    ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "idctr", "pro_controltardet", "N", xCon)
    GrabarOperacion xIdUsuario, 179, 7, xHorIni, Time, Date, xCon, ID_
    ' INGRESOS Y SALIDAS DE ALMACEN
    ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "id", "alm_ingreso", "N", xCon)
    GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, ID_
        
    cambiarEstadoRelacionados = True
    Exit Function
ERROR_:
    MsgBox "Ha ocurrido un error al tratar de cambiar de estado", vbInformation, xTitulo
    cambiarEstadoRelacionados = False
End Function

Private Function verificarCambioEstado(IDPROCORR_ As Double, ByRef MENSAJE_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
        
    ' Buscando Registros de Solicitud de Materiales
    cSQL = "SELECT pro_ordenproddet.idprocorr, pro_ordenproddet.estado " _
        + vbCr + "FROM pro_ordenproddet " _
        + vbCr + "WHERE (((pro_ordenproddet.idprocorr)=" & IDPROCORR_ & ") AND ((pro_ordenproddet.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Solicitud de Materiales"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    ' Buscando Registros de Planilla
    cSQL = "SELECT pro_controltardet.idprocorr, pro_controltardet.estado " _
        + vbCr + "FROM pro_controltardet " _
        + vbCr + "WHERE (((pro_controltardet.idprocorr)=" & IDPROCORR_ & ") AND ((pro_controltardet.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Planilla"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    ' Buscando Registros de Almacen
    cSQL = "SELECT alm_ingreso.idprocorr, alm_ingreso.estado " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idprocorr)=" & IDPROCORR_ & ") AND ((alm_ingreso.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Almacen"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    verificarCambioEstado = True
    Exit Function
    
SALIR_:
    MENSAJE_ = "Se han encontrado " & MENSAJE_ & " que se encuentran en un estado no modificable; " _
    & vbCr & "verifique la condición de dichos Registros para completar esta acción."
End Function

Private Sub Fg1_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Dim IDPROCORR_ As Double
    Dim ESTADO_ As Double
    Dim Rpta As Integer
    Dim MENSAJE_ As String
    
    If Col = COLUMNAESTADO_ Then
        Rpta = MsgBox("¿ Esta seguro de cambiar el estado actual?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        
        If Rpta = vbNo Then
            Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_) = ESTADOANTERIOR_
            Exit Sub
        End If
                
        IDPROCORR_ = NulosN(Fg1.TextMatrix(Row, COLUMNACORR_))
        ESTADO_ = NulosN(Fg1.TextMatrix(Row, Col))
        
        If ESTADOANTERIOR_ > ESTADO_ Then
            MsgBox "Este cambio de estado no esta permitido", vbInformation, xTitulo
        Else
            If ESTADO_ <> ESTADOANULADO_ Then Exit Sub
            
            If verificarCambioEstado(IDPROCORR_, MENSAJE_) Then
                If cambiarEstadoRelacionados(IDPROCORR_, ESTADO_) Then
                    CAMBIOGRABAR_ = -1
                End If
            Else
                MsgBox MENSAJE_, vbInformation, xTitulo
                Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_) = ESTADOANTERIOR_
            End If
        End If
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = COLUMNAUM_ Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col = Fg1.Cols - 1 Then
            Fg1.Editable = flexEDNone
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    Select Case Col
        Case COLUMNANUMPROD_, COLUMNATOTPROD_, COLUMNAHORINI_, COLUMNAHORFIN_
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        
        Case COLUMNAOBS_
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    
    If KeyCode = 117 Then
        Dg(TabOne2.CurrTab).SetFocus
    End If
    
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then ' F3 = Agregar Item
        cmd_Click 0
    End If
    
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then ' F4 = Eliminar Item
        cmd_Click 1
    End If
    Exit Sub
    
error:
    SHOW_ERROR Me.Name, "Fg1_KeyUp"
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then
            PopupMenu menu4
        Else
            PopupMenu menu1
        End If
    End If
End Sub

Private Sub Fg1_RowColChange()
    Dim N_FILTER As String
    If Agregando = True Then Exit Sub
    If Fg1.Rows = 1 Then
        Exit Sub
    End If
    
    If Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_) = "" Then ' NO HAY RECETA
        N_FILTER = "-999"
    Else
        N_FILTER = Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_)
    End If
    
    If Fg1.Row <= 0 Then Exit Sub
    Label2.Caption = Fg1.TextMatrix(Fg1.Row, COLUMNAPRODUCTO_)
    
    ' FILTRANDO LOS INSUMOS Y TAREAS
    RST_INSUMO.Filter = "idparte ='" + Fg1.TextMatrix(Fg1.Row, COLUMNAIDPARTE_) + "' AND idrec='" + N_FILTER + "'"
    RST_TAREA.Filter = "idparte ='" + Fg1.TextMatrix(Fg1.Row, COLUMNAIDPARTE_) + "' AND idrec='" + N_FILTER + "'"
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    
    '--Almacenar temporalmente el codigo del menu
    IdMenuActivo = xIdMenu
    
    mMesActivo = xMes
    
    pCargarGrid
    
    SeEjecuto = True

End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    '***************************
    iniciarCampos
    '***************************
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 5000 Then Me.Height = 5000

    ' Consulta
    TabOne1.Width = Me.Width - 90
    TabOne1.Height = Me.Height - 765
    Label4(0).Width = TabOne1.Width - 300
    Dg3.Width = TabOne1.Width - 165
    Dg3.Height = TabOne1.Height - 780
    lblperiodo(0).Left = TabOne1.Width - 2040
    
    ' Detalle
    Label1.Width = TabOne1.Width - 300
    Frame3.Left = TabOne1.Width - 1560
    Frame4.Left = TabOne1.Width - 6570
    lbl_cb(1).Width = TabOne1.Width - 8775
    Fg1.Width = TabOne1.Width - 210
    Fg1.Height = Int(TabOne1.Height / 2) - 1162
    Label2.Top = Fg1.Height + 1050
    TabOne2.Top = Label2.Top + 225
    TabOne2.Width = TabOne1.Width - 210
    TabOne2.Height = Int(TabOne1.Height / 2) - 562
    ' Insumos
    Dg(0).Width = TabOne2.Width - 255
    Dg(0).Height = TabOne2.Height - 570
    ' Tareas
    Dg(1).Width = TabOne2.Width - 255
    Dg(1).Height = TabOne2.Height - 570
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '******************************
    If CAMBIOGRABAR_ = -1 Then
        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
    End If
    '******************************
    
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    Else
        Set RstFrm = Nothing
    End If
End Sub

Private Sub Menu1_1_Click()
    pRegistroAdd
End Sub

Private Sub Menu1_2_Click()
    pRegistroDel
End Sub

Private Sub Menu1_3_Click()
    pRegistroAdd False, True
End Sub

Private Sub Menu1_4_Click()
    pRegistroAdd True, True
End Sub

Private Sub Menu2_1_Click()
    pRegistroDel
End Sub

Private Sub menu2_2_Click()
    Dim Q_ROW As Long
    If Fg1.Rows <= 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    Do While Fg1.Rows > 1
        Fg1.Row = 1
        pRegistroDel
    Loop
    Agregando = False
End Sub

Private Sub Menu3_1_Click()
    pRegistroAdd True, True
End Sub

Private Sub Menu3_2_Click()
    pRegistroAdd True, True
End Sub

Private Sub Menu3_3_Click()
    pRegistroAdd True, False
End Sub

Private Sub Menu4_1_1_Click()
    ' CONSULTA DE INSUMO/ X PRODUCTO
    CARGAR_FRM_PRODUCCION_LISTA E_INSUMO, 0, -1, Fg1.TextMatrix(Fg1.Row, COLUMNAIDREC_)
End Sub

Private Sub Menu4_1_2_Click()
    ' CONSULTA DE INSUMO/ TODOS LOS PRODUCTOS
    CARGAR_FRM_PRODUCCION_LISTA E_INSUMO, 1
End Sub

Private Sub Menu4_1_3_Click()
    ' CONSULTA DE INSUMO/ RESUMEN
    CARGAR_FRM_PRODUCCION_LISTA E_INSUMO, 2
End Sub

Private Sub Menu5_1_Click()
    ' AGREGAR
    Dg_KeyUp 0, 114, 0
End Sub

Private Sub Menu5_3_Click()
    ' ELIMINAR
    Dg_KeyUp 0, 115, 0
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
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        RstFrm.Filter = ""
        TDB_FiltroLimpiar Dg3
    End If
    
    If Button.Index = 10 Then CambiarMes
    
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 13 Then pExportar
    
    If Button.Index = 16 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_produccion
'* Parametros       :
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
    
    TabOne1.CurrTab = 0
    
    If MsgBox("¿Esta seguro de eliminar la Producción?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETE * FROM pro_producciondetins WHERE idpro = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM pro_producciondettar WHERE idpro = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM pro_producciondet WHERE idpro = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM pro_produccion WHERE id = " & RstFrm("id") & ""
        xCon.Execute "UPDATE pro_programadet SET idpro = 0 WHERE idpro = " & RstFrm("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "La Producción del dia " + Format(RstFrm("dia"), "dd/mm/yy") + " fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
        RstFrm.Requery
        Dg3.Refresh

    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O EDICION DE UN REGISTRO
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    If CAMBIOGRABAR_ = -1 Then
        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle Programa de Producción"
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    Dg3.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If

    QueHace = 2
    xHorIni = Time
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Habilitar_Obj True
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    Label1.Caption = "Modificando la Producción"
    GRID_COMBOLIST Fg1, COLUMNARECETA_
    GRID_COMBOLIST Fg1, COLUMNARESPONSABLE_
    GRID_COMBOLIST Fg1, COLUMNATURNO_
    GRID_COMBOLIST Fg1, COLUMNAORDENPROD_
   
    Fg1.ColEditMask(COLUMNAHORINI_) = "##:##"
    Fg1.ColEditMask(COLUMNAHORFIN_) = "##:##"
    Fg1.ColFormat(COLUMNANUMPROD_) = FORMAT_NUM_PARTE

    ' SI DESEA AGREGAR PRODUCTOS AL GRID OBTENER EL ULTIMO NUMERO DE PRODUCCION
    M_NUM_PARTE = HallaValor(xCon, "pro_producciondet", "numparte")
    
''    TxtFecha(0).Enabled = False
    txt_cb(1).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraSegundoTab()
    With RstFrm
        NUMEROPROD_ = HallaValor(xCon, "pro_producciondet", "numparte")
        TabOne2.CurrTab = 0
        Blanquea
                
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
        TxtFecha(0).valor = .Fields("dia") & ""
       
        txt_cb(1).Text = .Fields("supnum") & ""
        lbl_cb(1).Caption = .Fields("sup") & ""
        lbl_cb_cod(1).Caption = .Fields("idsup") & ""
        
        txt(0).Text = .Fields("id") & ""
        Fg1.ColFormat(COLUMNANUMPROD_) = FORMAT_NUM_PARTE
        MuestraDetalle
        If Fg1.Rows >= 2 Then Fg1.Row = 1:      Fg1.Col = COLUMNANUMPROD_:         Fg1_RowColChange
        
        llenarEstados
    End With
End Sub

Private Sub llenarEstados()
    Dim CAMPOS As String
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT * FROM mae_estados ORDER BY id"
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se ha encontrado estados, Ingrese estados", vbInformation, xTitulo
        Exit Sub
    End If
    
    xRs.MoveFirst
    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
    xRs.MoveNext
    While Not xRs.EOF
        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
        xRs.MoveNext
    Wend
    Fg1.ColComboList(COLUMNAESTADO_) = CAMPOS
End Sub

Private Sub iniciarCampos()
    Dim CAMPOS As String
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Rows = 1
    'LimpiarGrid Fg1
    
    COLUMNANUMPROD_ = 1
    COLUMNAPRODUCTO_ = 2
    COLUMNARECETA_ = 3
    COLUMNAUM_ = 4
    COLUMNAORDENPROD_ = 5
    COLUMNATOTPROG_ = 6
    COLUMNATOTPROD_ = 7
    COLUMNAHORINI_ = 8
    COLUMNAHORFIN_ = 9
    COLUMNARESPONSABLE_ = 10
    COLUMNATURNO_ = 11
    COLUMNAESTADO_ = 12
    '*****************
    COLUMNAOBS_ = 13
    '*****************
    COLUMNAIDPARTE_ = 14
    COLUMNAIDREC_ = 15
    COLUMNAIDITEM_ = 16
    COLUMNAIDUNID_ = 17
    COLUMNAIDRES_ = 18
    COLUMNAIDTURNO_ = 19
    COLUMNACORR_ = 20
    COLUMNAIDORD_ = 21
    
    Fg1.ColWidth(COLUMNAIDORD_) = 0
    Fg1.ColWidth(COLUMNATOTPROG_) = 0
    Fg1.ColWidth(COLUMNAIDPARTE_) = 0
    Fg1.ColWidth(COLUMNAIDREC_) = 0
    Fg1.ColWidth(COLUMNAIDITEM_) = 0
    Fg1.ColWidth(COLUMNAIDUNID_) = 0
    Fg1.ColWidth(COLUMNAIDRES_) = 0
    Fg1.ColWidth(COLUMNAIDTURNO_) = 0
    Fg1.ColWidth(COLUMNACORR_) = 0
    
    Fg1.ColWidth(COLUMNAHORINI_) = 0
    Fg1.ColWidth(COLUMNAHORFIN_) = 0
    Fg1.ColWidth(COLUMNARESPONSABLE_) = 0
    Fg1.ColWidth(COLUMNATURNO_) = 0
    
    Dg3.Columns("dia1").NumberFormat = FORMAT_DATE
        
    Dg(0).BatchUpdates = False
    Dg(1).BatchUpdates = False
    
    Dg(0).Columns("unid").NumberFormat = FORMAT_PU:
    Dg(0).Columns("canteo").NumberFormat = FORMAT_CANT:
    Dg(0).Columns("canprog").NumberFormat = FORMAT_CANT:
    Dg(0).Columns("canreal").NumberFormat = FORMAT_CANT:
    Dg(0).Columns("dif").NumberFormat = FORMAT_CANT:
    
    Dg3.Columns("desestado").Alignment = dbgCenter
    
    Dg(0).Columns("descripcion").Button = False
    Dg(0).Columns("descripcion").ButtonAlways = False

    Dg(1).Columns("horini").EditMask = "##:##"
    Dg(1).Columns("horfin").EditMask = "##:##"
    Dg(1).Columns("canper").NumberFormat = FORMAT_CANT
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.FrozenCols = 4
    Fg1.Tag = Fg1.FormatString
    Dg3.HeadLines = 2
    
    NUMEROPROD_ = 0
    
    llenarEstados
    
    CORR_ = -666
    CAMBIOGRABAR_ = 0
    ESTADOANTERIOR_ = 1
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim xCol, xFil As Integer
    Dim xSQL As String
    Dim xFch As Date
    Dim xFila  As Integer

    xSQL = "SELECT pro_producciondet.corr, pro_producciondet.numparte AS idparte, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, mae_turnos.id AS idturno, pro_producciondet.numparte, alm_inventario.descripcion AS proddesc, pro_receta.codrec, mae_unidades.abrev, pro_producciondet.canprog, pro_producciondet.cantidad AS canreal, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.idres, pla_empleados.nombre AS resnom, mae_turnos.descripcion AS turdesc, pro_producciondet.estado, pro_producciondet.obs, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numord, pro_producciondet.idord " _
        + vbCr + "FROM (pla_empleados RIGHT JOIN (alm_inventario INNER JOIN ((((mae_unidades RIGHT JOIN (pro_producciondet LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) ON mae_unidades.id = pro_producciondet.idunimed) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_programadet ON (pro_producciondet.idpro = pro_programadet.idpro) AND (pro_producciondet.idrec = pro_programadet.idrec)) LEFT JOIN mae_turnos ON pro_producciondet.idturno = mae_turnos.id) ON alm_inventario.id = pro_receta.iditem) ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((pro_producciondet.idpro) = " + CStr(RstFrm.Fields("id")) + ")) "
    
'    xSQL = "SELECT pro_producciondet.corr, pro_producciondet.numparte AS idparte, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, mae_turnos.id AS idturno, pro_producciondet.numparte, alm_inventario.descripcion AS proddesc, pro_receta.codrec, mae_unidades.abrev, pro_producciondet.canprog, pro_producciondet.numprog, pro_producciondet.cantidad AS canreal, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.idres, pla_empleados.nombre AS resnom, mae_turnos.descripcion AS turdesc, pro_producciondet.estado, pro_producciondet.obs " _
'        + vbCr + "FROM pla_empleados RIGHT JOIN (alm_inventario INNER JOIN ((((mae_unidades RIGHT JOIN (pro_producciondet LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) ON mae_unidades.id = pro_producciondet.idunimed) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_programadet ON (pro_producciondet.idrec = pro_programadet.idrec) AND (pro_producciondet.idpro = pro_programadet.idpro)) LEFT JOIN mae_turnos ON pro_producciondet.idturno = mae_turnos.id) ON alm_inventario.id = pro_receta.iditem) ON pla_empleados.id = pro_emp.idemp " _
'        + vbCr + "WHERE (((pro_producciondet.idpro) = " + CStr(RstFrm.Fields("id")) + ")) "

    RST_Busq xRs, xSQL, xCon
    If xRs.RecordCount <> 0 Then
        Agregando = True
        With Fg1
            .Rows = 1
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                xFila = .Rows
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, COLUMNANUMPROD_) = xRs.Fields("numparte") & ""
                .TextMatrix(.Rows - 1, COLUMNAPRODUCTO_) = xRs.Fields("proddesc") & ""
                .TextMatrix(.Rows - 1, COLUMNARECETA_) = xRs.Fields("codrec") & ""
                .TextMatrix(.Rows - 1, COLUMNAUM_) = xRs.Fields("abrev") & ""
                .TextMatrix(.Rows - 1, COLUMNATOTPROG_) = Format(NulosN(xRs.Fields("canprog")), FORMAT_CANTIDAD) ' CANTIDAD PROGRAMADA
                .TextMatrix(.Rows - 1, COLUMNATOTPROD_) = Format(NulosN(xRs.Fields("canreal")), FORMAT_CANTIDAD) ' CANTIDAD REAL
                ' DE LA HORAS
                If IsDate(xRs.Fields("horini")) = True Then .TextMatrix(.Rows - 1, COLUMNAHORINI_) = Format(xRs.Fields("horini"), FORMAT_HORA_SIN_SEGUNDO)
                If IsDate(xRs.Fields("horfin")) = True Then .TextMatrix(.Rows - 1, COLUMNAHORFIN_) = Format(xRs.Fields("horfin"), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, COLUMNARESPONSABLE_) = xRs.Fields("resnom") & ""                              ' ID DEL RESPONSABLE
                .TextMatrix(.Rows - 1, COLUMNATURNO_) = xRs.Fields("turdesc") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDPARTE_) = xRs.Fields("idparte") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDREC_) = xRs.Fields("idrec") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDITEM_) = xRs.Fields("iditem") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDUNID_) = xRs.Fields("idunimed") & ""
                .TextMatrix(.Rows - 1, COLUMNAIDRES_) = xRs.Fields("idres") & ""                              ' NOMBRE DEL RESPONSABLE
                .TextMatrix(.Rows - 1, COLUMNAIDTURNO_) = xRs.Fields("idturno") & ""
                .TextMatrix(.Rows - 1, COLUMNAORDENPROD_) = xRs.Fields("numord") & ""
                .TextMatrix(.Rows - 1, COLUMNATOTPROG_) = Format(xRs.Fields("canprog"), FORMAT_CANTIDAD) & ""
                .TextMatrix(.Rows - 1, COLUMNAOBS_) = NulosC(xRs("obs"))
                .TextMatrix(.Rows - 1, COLUMNAESTADO_) = NulosN(xRs.Fields("estado"))
                .TextMatrix(.Rows - 1, COLUMNACORR_) = NulosN(xRs.Fields("corr"))
                .TextMatrix(.Rows - 1, COLUMNAIDORD_) = NulosN(xRs.Fields("idord"))
                
                DATOS_TMP_ADD xRs.Fields("idparte") & "", xRs.Fields("idrec"), E_INSUMO
                DATOS_TMP_ADD xRs.Fields("idparte") & "", xRs.Fields("idrec"), e_TAREA
                
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
        End With
    End If
    
    Set xRs = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub

error:
    SHOW_ERROR Me.Name, "MuestraDetalle"
    Me.MousePointer = vbDefault
    Set xRs = Nothing
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : Habilitar_Obj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA LOS CONTROLES TEXTBOXY COMMAND DEL FORMULARIO
'* Parametros       : NOMBRE   |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band     |  Boolean    |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
    habilitar Cmd, band
    Dg(0).Splits(0).Locked = Not band
    Dg(1).Splits(0).Locked = Not band
    
    If band = True Then
        Dg(0).MarqueeStyle = dbgHighlightCell
        Dg(1).MarqueeStyle = dbgHighlightCell
        Dg(0).Columns("descripcion").Button = True
    Else
        Dg(0).MarqueeStyle = dbgHighlightRow
        Dg(1).MarqueeStyle = dbgHighlightRow
        Dg(0).Columns("descripcion").Button = False
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX PARA EL INGRESO DE DATOS
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    LimpiaText TxtFecha
    LimpiaText txt
    LimpiaText txt_cb
    
    Fg1.Rows = Fg1.FixedRows
    
    Set Dg(0).DataSource = Nothing
    Set Dg(1).DataSource = Nothing

    pDefinirRst RST_INSUMO, E_INSUMO
    pDefinirRst RST_TAREA, e_TAREA
    Set Dg(0).DataSource = RST_INSUMO
    Set Dg(1).DataSource = RST_TAREA
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Nuevo()
    Dim CAMPOS As String
    
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne2.CurrTab = 0
    TabOne1.TabEnabled(0) = False
    ActivaTool
    If mMesActivo <> 0 And mMesActivo <> 13 Then TxtFecha(0).valor = CDate("01/" & mMesActivo & "/" & AnoTra)
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregando Producción"
    GRID_COMBOLIST Fg1, COLUMNARECETA_
    GRID_COMBOLIST Fg1, COLUMNARESPONSABLE_
    GRID_COMBOLIST Fg1, COLUMNATURNO_
    
    '******************************************
    GRID_COMBOLIST Fg1, COLUMNAORDENPROD_
    '******************************************
    
    M_NUM_PARTE = HallaValor(xCon, "pro_producciondet", "numparte")
    Fg1.ColEditMask(COLUMNAHORINI_) = "##:##"
    Fg1.ColEditMask(COLUMNAHORFIN_) = "##:##"
    Fg1.ColFormat(COLUMNANUMPROD_) = FORMAT_NUM_PARTE
    
    llenarEstados
    
    TxtFecha(0).Enabled = True
    TxtFecha(0).valor = Date
    txt_cb(1).SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pImprimir True

    If ButtonMenu.Index = 2 Then pImprimir
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_produccion, ESTA FUNCION DEVUELEVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Parametros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Producción", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstIns As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xId As Double
    Dim xCodDet&    '--al detalle
    Dim xCol&, xFil&
    Dim xCorr&
    
On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    
    '***************************************************************
    Dim CORRAUX_ As Double
    CORRAUX_ = HallaCodigoTabla("pro_producciondet", xCon, "corr")
    '***************************************************************
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_produccion ", xCon
        xId = HallaCodigoTabla("pro_produccion", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pro_produccion WHERE id = " & xId & ";", xCon
        
        ' restar el stock actual encabezado
        RST_Busq RstTmp, "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS total  FROM pro_producciondet AS pro_producciondet GROUP BY pro_producciondet.idpro, pro_producciondet.iditem HAVING (((pro_producciondet.idpro)=" & xId & "));", xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = [alm_inventario].[stckact] - " & NulosN(RstTmp("total")) & " WHERE (((alm_inventario.id)=" & RstTmp("iditem") & "));"
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        
        ' acumular el stock actual detalle
        RST_Busq RstTmp, "SELECT pro_producciondetins.iditem, Sum(pro_producciondetins.canutil) AS total FROM pro_producciondet INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) GROUP BY pro_producciondetins.iditem, pro_producciondet.idpro HAVING (((pro_producciondet.idpro)=" & xId & "));", xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = [alm_inventario].[stckact] + " & NulosN(RstTmp("total")) & " WHERE (((alm_inventario.id)=" & RstTmp("iditem") & "));"
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        
        ' eliminando los registros
        xCon.Execute "DELETE * FROM pro_producciondetins WHERE idpro = " & xId & ""
        xCon.Execute "DELETE * FROM pro_producciondettar WHERE idpro = " & xId & ""
        xCon.Execute "DELETE * FROM pro_producciondet WHERE idpro = " & xId & ""
        
        '--actualizando el codigo de produccion del programa de produccion a 0
        xCon.Execute "UPDATE pro_programadet SET idpro =0 WHERE idpro = " & xId & ""
    End If
    
    mIdRegistro = xId
    
    RST_Busq RstDet, "SELECT top 1 * FROM pro_producciondet", xCon
    RST_Busq RstIns, "SELECT top 1 * FROM pro_producciondetins", xCon
    RST_Busq RstTar, "SELECT top 1 * FROM pro_producciondettar", xCon
    
    RstCab("dia") = CDate(TxtFecha(0).valor)
    RstCab("idsup") = NulosN(lbl_cb_cod(1).Caption)
    RstCab("num") = Format(xId, FORMAT_NUM_PRODUCCION)
    RstCab("obs") = ""
    
    RstCab.Update
    
    Dim F_CAMBIO_PRODUCCION As Boolean
    Dim M_PRODUCCION As String
        
    F_CAMBIO_PRODUCCION = False
    For xFil = 1 To Fg1.Rows - 1
        ' 1 = NUM PRODUCCION    12 = RECETA
        If NulosN(Fg1.TextMatrix(xFil, COLUMNANUMPROD_)) > 0 And NulosN(Fg1.TextMatrix(xFil, COLUMNAIDREC_)) > 0 Then
            RstDet.AddNew
            
            xCodDet = NulosN(Fg1.TextMatrix(xFil, COLUMNANUMPROD_))
            
            RstDet("idpro") = xId
            
            '******************************************************************
            ' Se llena el correlativo
            If NulosN(Fg1.TextMatrix(xFil, COLUMNACORR_)) < 0 Then
                RstDet("corr") = CORRAUX_
                CORRAUX_ = CORRAUX_ + 1
            Else
                RstDet("corr") = NulosN(Fg1.TextMatrix(xFil, COLUMNACORR_))
            End If
            '******************************************************************
            
            ' VALIDAR QUE EL NUMERO DE PRODUCCION SEA DIFERENTE
            M_PRODUCCION = Format(xCodDet, FORMAT_NUM_PARTE)
            Set RstTmp = Nothing
            
            If QueHace = 1 Then
                RST_Busq RstTmp, "SELECT pro_producciondet.numparte FROM pro_producciondet WHERE (((pro_producciondet.numparte)='" + M_PRODUCCION + "') AND ((pro_producciondet.idpro)<>" + CStr(xId) + "));", xCon
                If RstTmp.EOF = False Or RstTmp.BOF = False Or RstTmp.RecordCount <> 0 And xFil = 1 Then
                    M_PRODUCCION = Format(HallaValor(xCon, "pro_producciondet", "numparte"), FORMAT_NUM_PARTE)
                    F_CAMBIO_PRODUCCION = True
                End If
                Set RstTmp = Nothing
            End If
            
            RstDet("numparte") = M_PRODUCCION
            RstDet("idrec") = Fg1.TextMatrix(xFil, COLUMNAIDREC_)
            ' FIN
            
            RstDet("iditem") = Fg1.TextMatrix(xFil, COLUMNAIDITEM_)
            RstDet("idunimed") = NulosN(Fg1.TextMatrix(xFil, COLUMNAIDUNID_))
            RstDet("cantidad") = NulosN(Fg1.TextMatrix(xFil, COLUMNATOTPROD_))
            If (Fg1.TextMatrix(xFil, COLUMNAHORINI_) <> "") Then RstDet("horini") = Fg1.TextMatrix(xFil, COLUMNAHORINI_)
            If (Fg1.TextMatrix(xFil, COLUMNAHORFIN_) <> "") Then RstDet("horfin") = Fg1.TextMatrix(xFil, COLUMNAHORFIN_)
            If (Fg1.TextMatrix(xFil, COLUMNAIDRES_) <> "") Then RstDet("idres") = NulosN(Fg1.TextMatrix(xFil, COLUMNAIDRES_))
            If (Fg1.TextMatrix(xFil, COLUMNAIDTURNO_) <> "") Then RstDet("idturno") = NulosN(Fg1.TextMatrix(xFil, COLUMNAIDTURNO_))
            RstDet("canprog") = NulosN(Fg1.TextMatrix(xFil, COLUMNATOTPROG_))
            RstDet("numprog") = NulosC(Fg1.TextMatrix(xFil, COLUMNAORDENPROD_))
            RstDet("estado") = NulosN(Fg1.TextMatrix(xFil, COLUMNAESTADO_))
            RstDet("obs") = NulosC(Fg1.TextMatrix(xFil, COLUMNAOBS_))
            
            ' -----------------------------ORDEN DE PRODUCCION
            RstDet("idord") = NulosN(Fg1.TextMatrix(xFil, COLUMNAIDORD_))
            
            RstDet.Update
            ' ACTUALIZAR EL PROGRAMA DE ACUERDO A LA FECHA Y RECETA
            xCon.Execute "UPDATE pro_programadet SET idpro =" + CStr(xId) + " WHERE dia = CDATE('" + TxtFecha(0).valor + "') AND idrec = " + CStr(Fg1.TextMatrix(xFil, COLUMNAIDREC_)) + ";"
            
            ' ADD INSUMOS
            RST_INSUMO.Filter = "idparte = " + Fg1.TextMatrix(xFil, COLUMNAIDPARTE_) + " AND idrec=" + Fg1.TextMatrix(xFil, COLUMNAIDREC_)
            
            If RST_INSUMO.RecordCount > 0 Then
                RST_INSUMO.MoveFirst
                Do While Not RST_INSUMO.EOF
                    If NulosN(RST_INSUMO.Fields("iditem")) <> 0 Then
                        RstIns.AddNew
                        ' CLAVE
                        RstIns("idpro") = xId
                        RstIns("numparte") = M_PRODUCCION
                        RstIns("idrec") = Fg1.TextMatrix(xFil, COLUMNAIDREC_)
                        RstIns("iditem") = NulosN(RST_INSUMO.Fields("iditem"))
                        ' FIN CLAVE
                        
                        RstIns("idunimed") = NulosN(RST_INSUMO.Fields("idunimed"))
                        RstIns("canutil") = NulosN(RST_INSUMO.Fields("canreal"))
                        RstIns("canpro") = NulosN(RST_INSUMO.Fields("unid"))
                        RstIns.Update
                    End If
                    RST_INSUMO.MoveNext
                Loop
            End If
            
            ' ADD TAREAS
            RST_TAREA.Filter = "idparte = " + Fg1.TextMatrix(xFil, COLUMNAIDPARTE_) + " AND idrec=" + Fg1.TextMatrix(xFil, COLUMNAIDREC_)
            xCorr = 1
            If RST_TAREA.RecordCount > 0 Then
                RST_TAREA.MoveFirst
                Do While Not RST_TAREA.EOF
                    RstTar.AddNew
                    '--CLAVE
                    RstTar("idpro") = xId
                    RstTar("numparte") = M_PRODUCCION
                    RstTar("idrec") = Fg1.TextMatrix(xFil, COLUMNAIDREC_)
                    RstTar("idtar") = NulosN(RST_TAREA.Fields("idtar"))
                    
                    RstTar("corr") = xCorr
                    '-FIN CLAVE
                    
                    RstTar("idunimed") = NulosN(RST_TAREA.Fields("idunimed"))

                    If IsDate(RST_TAREA.Fields("horini")) = True Then RstTar("horini") = CDate(RST_TAREA.Fields("horini"))
                    If IsDate(RST_TAREA.Fields("horfin")) = True Then RstTar("horfin") = CDate(RST_TAREA.Fields("horfin"))
                    RstTar("canper") = NulosN(RST_TAREA.Fields("canper"))
                    RstTar.Update
                    RST_TAREA.MoveNext
                    xCorr = xCorr + 1
                Loop
            End If
        End If
    Next xFil
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    MsgBox "La Producción se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + _
    IIf(F_CAMBIO_PRODUCCION = True, vbCr + "El número de Producción se cambió", ""), vbInformation, xTitulo
    Grabar = True
    CAMBIOGRABAR_ = 0

SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    Exit Function

LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS, ESTA FUNCION DEVUELVE
'*                    VERDADERO SI LOS DATOS SON CORRECTOS
'* Parametros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If TxtFecha(0).valor = "" Or IsDate(TxtFecha(0).valor) = False Then
        MsgBox "No ha especificado la fecha de Producción ", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    Dim Q_ROW  As Long
    Dim Q_COL As Long       ' COLUMNA A POSICIONAR SI FALTAN DATOS
    Dim band As Integer
    band = Validar(txt_cb)
    
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
       txt_cb(band).SetFocus
       Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado los productos para la producción", vbInformation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    ' VALIDAR EL INGRESO DE LOS DATOS
    Q_COL = -1
    For Q_ROW = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(Q_ROW, COLUMNANUMPROD_)) = 0 Then
            MsgBox "Ingrese El número de Producción del Producto:" + vbCr + _
            "Producto:        " + Fg1.TextMatrix(Q_ROW, COLUMNAPRODUCTO_) & "", vbExclamation, xTitulo
            Q_COL = COLUMNANUMPROD_:          Exit For
        ElseIf NulosN(Fg1.TextMatrix(Q_ROW, COLUMNAIDREC_)) = 0 Then
            MsgBox "Ingrese La Receta del Producto:" + vbCr + _
            "Producto:        " + Fg1.TextMatrix(Q_ROW, COLUMNAPRODUCTO_) & "", vbExclamation, xTitulo
            Q_COL = COLUMNAIDREC_:          Exit For
        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, COLUMNATOTPROD_)) = False Or Fg1.TextMatrix(Q_ROW, COLUMNATOTPROD_) = "0" Then
            MsgBox "Ingrese el total de Producción:" + vbCr + _
            "Producto:        " + Fg1.TextMatrix(Q_ROW, COLUMNAPRODUCTO_) & "" + vbCr + _
            "Receta:         " + Fg1.TextMatrix(Q_ROW, COLUMNARECETA_) & "" + vbCr, vbExclamation, xTitulo
            Q_COL = COLUMNATOTPROD_:          Exit For
'        ElseIf IsDate(Fg1.TextMatrix(Q_ROW, COLUMNAHORINI_)) = False Or IsDate(Fg1.TextMatrix(Q_ROW, COLUMNAHORFIN_)) = False Then
'            MsgBox "Ingrese el la Hora " + IIf(IsDate(Fg1.TextMatrix(Q_ROW, COLUMNAHORINI_)) = False, "Inicial", "Final") + " de la Producción" + vbCr + _
'            "Producto:  " + Fg1.TextMatrix(Q_ROW, COLUMNAPRODUCTO_) & "" + vbCr + _
'            "Receta:         " + Fg1.TextMatrix(Q_ROW, COLUMNARECETA_) & "" + vbCr, vbExclamation, xTitulo
'            Q_COL = IIf(IsDate(Fg1.TextMatrix(Q_ROW, COLUMNAHORINI_)) = False, COLUMNAHORINI_, COLUMNAHORFIN_):        Exit For
'        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, COLUMNAIDRES_)) = False Or Fg1.TextMatrix(Q_ROW, COLUMNAIDRES_) = "0" Then
'            MsgBox "Ingrese el Responsable de la Producción:" + vbCr + _
'            "Producto:  " + Fg1.TextMatrix(Q_ROW, COLUMNAPRODUCTO_) & "" + vbCr + _
'            "Receta:         " + Fg1.TextMatrix(Q_ROW, COLUMNARECETA_) & "" + vbCr, vbExclamation, xTitulo
'            Q_COL = COLUMNARESPONSABLE_:       Exit For
'        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, COLUMNAIDTURNO_)) = False Or Fg1.TextMatrix(Q_ROW, COLUMNAIDTURNO_) = "0" Then
'            MsgBox "Ingrese el Turno de la Producción:" + vbCr + _
'            "Producto:  " + Fg1.TextMatrix(Q_ROW, COLUMNAPRODUCTO_) & "" + vbCr + _
'            "Receta:         " + Fg1.TextMatrix(Q_ROW, COLUMNARECETA_) & "" + vbCr, vbExclamation, xTitulo
'            Q_COL = COLUMNATURNO_:       Exit For
        End If
    Next Q_ROW
    
    If Q_COL <> -1 Then
        Agregando = True:  Fg1.Row = Q_ROW: Fg1.Col = Q_COL: Agregando = False
        Exit Function
    End If

    fValidarDatos = True
End Function

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA TABLA pro_produccion EN EL CONTROL Dg3
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    Dim xSQL  As String
    
    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = lblperiodo(0).Caption
    
    TDB_FiltroLimpiar Dg3
    Set RstFrm = Nothing
    
    '------------------------------------------------------------------------------------------
    '--Bloqueamos los botones del Toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    xSQL = "SELECT pro_produccion.id, pro_produccion.num, pro_produccion.dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, [pla_empleados_1].[apepat] & ' ' & [pla_empleados_1].[apemat] & ' ' & [pla_empleados_1].[nom] AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, [pla_empleados_2].[apepat] & ' ' & [pla_empleados_2].[apemat] & ' ' & [pla_empleados_2].[nom] AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec, pro_producciondet.cantidad & '' AS cantidad1, pro_produccion.dia & '' AS dia1, mae_unidades.abrev, Format([pro_producciondet].[horini],'Short Time') AS horiniF, Format([pro_producciondet].[horfin],'Short Time') AS horfinF, pro_producciondet.estado AS idestado, UCase([mae_estados].[descripcion]) AS desestado, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numordprod " _
        + vbCr + "FROM (((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN ((((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN mae_estados ON pro_producciondet.estado = mae_estados.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND Month(pro_produccion.dia)= " + CStr(mMesActivo) + " " _
        + vbCr + "ORDER BY pro_produccion.dia, pro_produccion.num;"

'    xSQL = "SELECT pro_produccion.id, pro_produccion.num, pro_produccion.dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, [pla_empleados_1].[apepat] & ' ' & [pla_empleados_1].[apemat] & ' ' & [pla_empleados_1].[nom] AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, [pla_empleados_2].[apepat] & ' ' & [pla_empleados_2].[apemat] & ' ' & [pla_empleados_2].[nom] AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec, pro_producciondet.cantidad & '' AS cantidad1, pro_produccion.dia & '' AS dia1, mae_unidades.abrev, Format([horini],'Short Time') AS horiniF, Format([horfin],'Short Time') AS horfinF, pro_producciondet.estado AS idestado, UCase([mae_estados].[descripcion]) AS desestado " _
'        + vbCr + "FROM (((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN (((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN mae_estados ON pro_producciondet.estado = mae_estados.id " _
'        + vbCr + "WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND MONTH(pro_produccion.dia)= " + CStr(mMesActivo) + " " _
'        + vbCr + "ORDER BY pro_produccion.dia, pro_produccion.num;"

'    xSQL = "SELECT pro_produccion.id, pro_produccion.num, pro_produccion.dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, [pla_empleados_1].[apepat] & ' ' & [pla_empleados_1].[apemat] & ' ' & [pla_empleados_1].[nom] AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, [pla_empleados_2].[apepat] & ' ' & [pla_empleados_2].[apemat] & ' ' & [pla_empleados_2].[nom] AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec, pro_producciondet.cantidad & '' AS cantidad1, pro_produccion.dia & '' AS dia1, pro_producciondet.numprog, pro_producciondet.canprog, [pro_producciondet].[canprog]-[pro_producciondet].[cantidad] AS desv, IIf([pro_producciondet].[canprog]=0,'',([desv]/[pro_producciondet].[canprog])*100) AS desvporc, mae_unidades.abrev, Format([horini],'Short Time') AS horiniF, Format([horfin],'Short Time') AS horfinF, pro_producciondet.estado AS idestado, UCase([mae_estados].[descripcion]) AS desestado " _
'        + vbCr + "FROM (((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN (((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN mae_estados ON pro_producciondet.estado = mae_estados.id " _
'        + vbCr + "WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND MONTH(pro_produccion.dia)= " + CStr(mMesActivo) + " " _
'        + vbCr + "ORDER BY pro_produccion.dia, pro_produccion.num;"
    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, xSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

'*****************************************************************************************************
'* Nombre           : CambiarMes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL MES DE TRABAJO
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Selecione un Periodo Correcto", vbExclamation, xTitulo
        CambiarMes
        Exit Sub
    End If
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : pDefinirRst
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DEFINIR EL RECORSET TEMPORAL PARA INSUMO Y TAREA
'* Parametros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    rst       |  ADODB.Recordset  |
'*                    fTipo     |  e_PROGRAMA       |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pDefinirRst(Rst As ADODB.Recordset, fTipo As e_PROGRAMA)
    Dim RST_ORIGEN As New ADODB.Recordset
    Dim nSQL As String
    nSQL = pGenerarConsulta("-1", "-1", fTipo, CDate("01/01/07"), True)
    RST_Busq RST_ORIGEN, nSQL, xCon
    DEFINIR_RST_TMP Rst, RST_ORIGEN
    Set RST_ORIGEN = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS DATOS DEL CONTROL Dg3
'* Parametros       : NOMBRE       |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IMP_LISTADO  |  Boolean    |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir(Optional IMP_LISTADO As Boolean = False)
    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE PRODUCCIÓN", "LISTADO DE PRODUCCIÓN  -  Periodo: " + MonthName(mMesActivo, False)
    End If

    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

'*****************************************************************************************************
'* Nombre           : CARGAR_FRM_PRODUCCION_LISTA
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmManProduccion_lista
'* Parametros       : NOMBRE       |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    VENTANA      |  e_PROGRAMA   |
'*                    ESTILO_VISTA |  Integer      |
'*                    ID_PARTE     |  String       |
'*                    ID_RECETA    |  String       |
'* Devuelve         :
'*****************************************************************************************************
Private Sub CARGAR_FRM_PRODUCCION_LISTA(VENTANA As e_PROGRAMA, _
                                        ESTILO_VISTA As Integer, _
                                        Optional ID_PARTE As String = "-1", _
                                        Optional ID_RECETA As String = "-1")
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then Exit Sub
   
    With FrmManProduccion_lista
        .RECIBE_LINK_FRM CStr(RstFrm.Fields("id")), ID_PARTE, ID_RECETA, VENTANA, ESTILO_VISTA, Format(TxtFecha(0).valor, FORMAT_DATE), Trim(lbl_cb(1).Caption)
        .Show
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET xRs
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Fch.Prod":         xCampos(0, 1) = "dia":       xCampos(0, 2) = "850":    xCampos(0, 3) = "F"
    xCampos(1, 0) = "N°.Prod":          xCampos(1, 1) = "num":       xCampos(1, 2) = "900":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Receta":           xCampos(2, 1) = "codrec":    xCampos(2, 2) = "1000":   xCampos(2, 3) = "C"
    xCampos(3, 0) = "Producto":         xCampos(3, 1) = "proddesc":  xCampos(3, 2) = "3200":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cantidad":         xCampos(4, 1) = "cantidad":  xCampos(4, 2) = "800":    xCampos(4, 3) = "N"
    xCampos(5, 0) = "Responsable":      xCampos(5, 1) = "resnom":    xCampos(5, 2) = "1500":   xCampos(5, 3) = "C"
        
    nSQL = " SELECT pro_produccion.id, pro_produccion.num, format(pro_produccion.dia,'dd/mm/yy') as dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, [pla_empleados_1].[apepat] & ' ' & [pla_empleados_1].[apemat] & ' ' & [pla_empleados_1].[nom] AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, [pla_empleados_2].[apepat] & ' ' & [pla_empleados_2].[apemat] & ' ' & [pla_empleados_2].[nom] AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec " _
    + vbCr + " FROM ((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN ((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) ON pro_produccion.id = pro_producciondet.idpro " _
    + vbCr + " WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND MONTH(pro_produccion.dia)= " + CStr(mMesActivo) + ""

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Producción", "dia", "proddesc", Principio
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
'* Descripcion      : EJECUTA UN FILTRO EN EL RECORDSET RSTFRM
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Filtrar()
    Dim xCampos(5, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Producto":     xCampos(0, 1) = "proddesc": xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Fch. Pro":     xCampos(1, 1) = "dia":      xCampos(1, 2) = "F":         xCampos(1, 3) = "1000"
    xCampos(2, 0) = "N° Prod.":     xCampos(2, 1) = "num":      xCampos(2, 2) = "C":         xCampos(2, 3) = "800"
    xCampos(3, 0) = "Receta":       xCampos(3, 1) = "codrec":   xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Cantidad":     xCampos(4, 1) = "cantidad": xCampos(4, 2) = "N":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Responsable":  xCampos(5, 1) = "resnom":   xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : HallaValor
'* Tipo             : FUNCION
'* Descripcion      :
'* Parametros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    conn      |  ADODB.Connection |
'*                    tabla     |  String           |
'*                    campo     |  String           |
'* Devuelve         : Long
'*****************************************************************************************************
Private Function HallaValor(conn As ADODB.Connection, tabla As String, campo As String) As Long
    Dim xRs As New ADODB.Recordset
    On Error GoTo error
    RST_Busq xRs, "SELECT top 1 CLng([" + campo + "]) AS num FROM " + tabla + " ORDER BY CLng([" + campo + "]) DESC;", conn
    If xRs.State = 1 Then
        If xRs.EOF = False And xRs.BOF = False And xRs.RecordCount <> 0 Then
            HallaValor = NulosN(xRs.Fields(0)) + 1
        Else
            HallaValor = 1
        End If
    Else
        HallaValor = -1
    End If
    Set xRs = Nothing
    Exit Function

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "HallarValor"
End Function

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txt_cb(1).Text) <> "" Then
            Cmd(0).SetFocus
        Else
            SendKeys vbTab
        End If
        Exit Sub
    End If
    
    Select Case Index
        Case 3: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else: If validar_numero(KeyAscii) = True And KeyAscii = 46 Then KeyAscii = 0
    End Select
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String

    nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id " _
        + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
        + vbCr + " Where pro_empdet.idfun = 2 and pla_empleados.numdoc ='" + Trim(txt_cb(Index).Text) + "'" _
        + vbCr + " ORDER BY pla_empleados.apepat; "
       
    If Index = 1 Then
        nSQL = Replace(nSQL, "pro_empdet.idfun = 2", "pro_empdet.idfun = 1")
    End If

    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index) = RST_TMP.Fields(0) & ""                ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & ""        ' NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & ""    ' CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    Set RST_TMP = Nothing
    Exit Sub

error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    ReDim xCampos(2, 3) As String
    On Error GoTo error
    
    nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id " _
            + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
            + vbCr + " Where pro_empdet.idfun = 2 " _
            + vbCr + " ORDER BY pla_empleados.apepat; "
       
    nTitulo = "Buscando Programadores "
    
    If Index = 1 Then
        nTitulo = "Buscando Supervisores"
        nSQL = Replace(nSQL, "pro_empdet.idfun = 2", "pro_empdet.idfun = 1")
    End If
    
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "DNI":      xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    txt_cb(Index) = xRs.Fields(0) & ""             ' TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & ""     ' NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" ' CODIGO

SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA MS EXCEL LOS DATOS DEL RECORDSET RstTmp
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(6, 3) As String
    
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Fch Prod":                     xCampos(0, 1) = "dia":          xCampos(0, 2) = 1:  xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Nº Prod":                      xCampos(1, 1) = "numparte":     xCampos(1, 2) = 1:  xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Receta":                       xCampos(2, 1) = "codrec":       xCampos(2, 2) = 0:  xCampos(2, 3) = "1000"
    xCampos(3, 0) = "Producto":                     xCampos(3, 1) = "proddesc":     xCampos(3, 2) = 0:  xCampos(3, 3) = "4500"
    xCampos(4, 0) = "Responsable de Producción":    xCampos(4, 1) = "resnom":       xCampos(4, 2) = 0:  xCampos(4, 3) = "4500"
    xCampos(5, 0) = "Cant. Prod.":                  xCampos(5, 1) = "cantidad":     xCampos(5, 2) = 2:  xCampos(5, 3) = "1050"
    xCampos(6, 0) = "Estado":                       xCampos(6, 1) = "desestado":    xCampos(6, 2) = 2:  xCampos(6, 3) = "1050"
    '**********************************************************************************************************************************
        
    Set RstTmp = RstFrm
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Producción", "Periodo: " & lblperiodo(0).Caption & "  -  " & AnoTra, "", "Listado de Producción", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub
