VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmTransfAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacén - Transferencias"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   1590
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8265
      Top             =   30
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
            Picture         =   "FrmTransfAlmacen.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransfAlmacen.frx":277E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7020
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12382
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   45
         TabIndex        =   19
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6090
            Left            =   30
            TabIndex        =   20
            Top             =   480
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   10742
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
            Columns(1).Caption=   "Fch. Mov."
            Columns(1).DataField=   "fchdoc"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Responsable"
            Columns(3).DataField=   "responsable"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Alm.Origen"
            Columns(4).DataField=   "almorigen"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Alm.Destino"
            Columns(5).DataField=   "almdestino"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2646"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2566"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=5265"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=5186"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=4974"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4895"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=131588"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=4815"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=4736"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            HeadLines       =   1,5
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
            _StyleDefs(24)  =   "Splits(0).Style:id=75,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=76,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=77,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=78,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=80,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=79,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=75"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=76"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=77"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=75"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=76"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=77"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=79"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=75"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=76"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=77"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=79"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=20,.parent=75"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=76"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=77"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=79"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=118,.parent=75,.alignment=3"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=76,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=77,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=79"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=75"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=76"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=77"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=79"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
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
            TabIndex        =   21
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta de Transferencias"
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
            Left            =   105
            TabIndex        =   22
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   12525
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame5 
            Caption         =   "[ Detalles ]"
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
            Height          =   4155
            Left            =   120
            TabIndex        =   39
            Top             =   2400
            Width           =   11520
            Begin VB.Frame Frame4 
               Height          =   3795
               Left            =   9880
               TabIndex        =   41
               Top             =   240
               Width           =   1530
               Begin VB.CommandButton Cmd 
                  Caption         =   "Agregar Item"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   3
                  Left            =   90
                  TabIndex        =   13
                  Top             =   180
                  Width           =   1305
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Eliminar Item"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   4
                  Left            =   90
                  TabIndex        =   14
                  Top             =   550
                  Width           =   1305
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   3675
               Left            =   120
               TabIndex        =   40
               Top             =   360
               Width           =   9675
               _cx             =   17066
               _cy             =   6482
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
               Rows            =   2
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmTransfAlmacen.frx":2B10
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
            Caption         =   "[ Destino ]"
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
            Height          =   1155
            Left            =   5880
            TabIndex        =   34
            Top             =   1185
            Width           =   5760
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   2
               Left            =   1715
               Picture         =   "FrmTransfAlmacen.frx":2C63
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   390
               Width           =   240
            End
            Begin VB.TextBox NumSerDestinoText 
               Height          =   300
               Left            =   1055
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   11
               Text            =   "NumS"
               Top             =   720
               Width           =   915
            End
            Begin VB.TextBox NumDocDestinoText 
               Height          =   300
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   12
               Text            =   "NumDocDest"
               Top             =   720
               Width           =   3450
            End
            Begin VB.TextBox IdAlmDestinoText 
               Height          =   300
               Left            =   1055
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   9
               Text            =   "IdAlmDestinoText"
               Top             =   360
               Width           =   915
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidalm"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   10890
               TabIndex        =   38
               Top             =   105
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Almacén"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   37
               Top             =   405
               Width           =   615
            End
            Begin VB.Label AlmDestinoLabel 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AlmDestinoLabel"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   2010
               TabIndex        =   36
               Top             =   360
               Width           =   3615
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H80000001&
               BackStyle       =   1  'Opaque
               Height          =   90
               Left            =   2010
               Top             =   840
               Width           =   105
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Num. Doc."
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   35
               Top             =   765
               Width           =   765
            End
         End
         Begin VB.Frame FrmReceta 
            Caption         =   "[ Origen ]"
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
            Height          =   1155
            Left            =   120
            TabIndex        =   29
            Top             =   1185
            Width           =   5760
            Begin VB.TextBox NumDocOrigenText 
               Height          =   300
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   8
               Text            =   "NumDocOrig"
               Top             =   720
               Width           =   3450
            End
            Begin VB.TextBox NumSerOrigenText 
               Height          =   300
               Left            =   1055
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   7
               Text            =   "NumS"
               Top             =   720
               Width           =   915
            End
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   1680
               Picture         =   "FrmTransfAlmacen.frx":2D95
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   390
               Width           =   240
            End
            Begin VB.TextBox IdAlmOrigenText 
               Height          =   300
               Left            =   1055
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   5
               Text            =   "IdAlmOrigen"
               Top             =   360
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Num. Doc."
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   33
               Top             =   765
               Width           =   765
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H80000001&
               BackStyle       =   1  'Opaque
               Height          =   90
               Left            =   2010
               Top             =   840
               Width           =   105
            End
            Begin VB.Label AlmOrigenLabel 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AlmOrigenLabel"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   2010
               TabIndex        =   32
               Top             =   360
               Width           =   3615
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Almacén"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   31
               Top             =   405
               Width           =   615
            End
            Begin VB.Label lblidalm 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidalm"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   10890
               TabIndex        =   30
               Top             =   105
               Visible         =   0   'False
               Width           =   510
            End
         End
         Begin VB.TextBox GlosaText 
            Height          =   300
            Left            =   7050
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "GlosaText"
            Top             =   750
            Width           =   4620
         End
         Begin VB.CommandButton Cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   7700
            Picture         =   "FrmTransfAlmacen.frx":2EC7
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   480
            Width           =   240
         End
         Begin VB.TextBox NumSerText 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   1
            Text            =   "NumS"
            Top             =   750
            Width           =   915
         End
         Begin VB.TextBox NumDocText 
            Height          =   300
            Left            =   2265
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "NumDocText"
            Top             =   750
            Width           =   1440
         End
         Begin AspaTextBoxFecha.TextBoxFecha FechaText 
            Height          =   300
            Left            =   1170
            TabIndex        =   0
            Top             =   450
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
            Valor           =   "18/09/2007"
         End
         Begin VB.TextBox IdResponsableText 
            Height          =   300
            Left            =   7050
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "IdResponsableText"
            Top             =   450
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   9
            Left            =   6030
            TabIndex        =   28
            Top             =   795
            Width           =   405
         End
         Begin VB.Label ResponsableLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ResponsableLabel"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7995
            TabIndex        =   26
            Top             =   450
            Width           =   3660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Index           =   2
            Left            =   6000
            TabIndex        =   25
            Top             =   480
            Width           =   930
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2115
            Top             =   870
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Num. Doc."
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   23
            Top             =   795
            Width           =   765
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Transferencia"
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
            TabIndex        =   18
            Top             =   30
            Width           =   11670
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Mov."
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   17
            Top             =   480
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
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
            Object.ToolTipText     =   "Agregar "
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageKey        =   "IMG7"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageKey        =   "IMG13"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "IMG11"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar                "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmTransfAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMINGRESOALMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO DE DOCUMENTOS NO CONTABLES DE INGRESO O SALIDA,
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 17/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstIng As New ADODB.Recordset                    ' RECORDSET PRINCIPAL QUE CARGARA TODAS LAS OPERACIONES REGISTRADAS
Dim QueHace As Integer                               ' VARIABLE QUE INDICA EL ESTADO DEL FORMULARIO 1 = NUEVO, 2 = MODIFICA; 3 = SOLO LECTURA
Dim SeEjecuto As Boolean                             ' VARIABLE UTILIZADA PARA EJECUTAR UNA SOLA VEZ EL EVENTO ACTIVATE
Dim Agregando As Boolean                             ' VARIABLE QUE INFORMA A LOS CONTROLES FlexGrid QUE SE ESTA AGREGADO UNA FILA
Dim Mostrando As Boolean
Dim CaracteresNumericos As String                    ' ESPECIFICA LOS CARACTERES NUMERICOS QUE PODRA SOPORTAR LOS CONTROLES TextBox
Dim CaracteresNumericos2 As String, vStr As String   ' ESPECIFICA LOS CARACTERES NUMERICOS QUE PODRA SOPORTAR LOS CONTROLES TextBox
Dim mIdRegistro&                                     ' identificador del registro
Dim fOrdenLista As Boolean                           ' especfica el orden de la lista de la consulta
Dim xHorIni As Date                                  ' ESPECIFICA LA HORA DE INICIO
Dim mMesActivo As Integer                            ' --indica el mes activo
Dim fCierrePeriodo As Boolean                        ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer                          ' INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String

Private Sub cmd_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim nSQLId As String
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
        
    If QueHace = 3 Then Exit Sub
    
    Select Case Index
        Case 0 ' Responsable
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                      
            
            nTitulo = "Buscando Responsables"
            
            cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
                + vbCr + "FROM pla_empleados " _
                + vbCr + "ORDER BY pla_empleados.nombre;"
                        
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "apenom", "apenom", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdResponsableText.Text = NulosN(xRs("id"))
            ResponsableLabel.Caption = NulosC(xRs("apenom"))
            GlosaText.SetFocus
            
            Set xRs = Nothing
        
        Case 1 ' Almacen Origen
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            cSQL = "SELECT alm_almacenes.* FROM alm_almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdAlmOrigenText.Text = NulosN(xRs("id"))
            AlmOrigenLabel.Caption = UCase(NulosC(xRs("descripcion")))
            NumSerOrigenText.SetFocus
            Set xRs = Nothing
            
        Case 2 ' Almacen Destino
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            cSQL = "SELECT alm_almacenes.* FROM alm_almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdAlmDestinoText.Text = NulosN(xRs("id"))
            AlmDestinoLabel.Caption = UCase(NulosC(xRs("descripcion")))
            NumSerDestinoText.SetFocus
            Set xRs = Nothing
            
        Case 3 ' Agregar Item
            AgregarItem
        
        Case 4 ' Eliminar Item
            EliminarItem
        
    End Select
End Sub

Sub AgregarItem()
    Dim xCampos(3, 4) As String
    Dim nSQLId As String
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    
    xCampos(0, 0) = "Ítem":        xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
    
    nTitulo = "Buscando Ítems"
    
    nSQLId = GENERAR_SQL_ID(Fg1, Fg1.ColIndex("IDITEM"), " AND alm_inventario.id", "NOT IN", True)
    
    cSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
        + vbCr + "FROM alm_inventario INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((alm_inventario.activo)=-1)) " & nSQLId
    
    Set xRs = Nothing
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                    "descripcion", "descripcion", Principio, ""
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("ITEM")) = UCase(NulosC(xRs("descripcion")))
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("IDITEM")) = NulosN(xRs("iditem"))
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("CODIGO")) = Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon)
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("IDUNIMED")) = Busca_Codigo(NulosN(xRs("iditem")), "id", "idunimed", "alm_inventario", "N", xCon)
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("UM")) = Busca_Codigo(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("IDUNIMED"))), "id", "abrev", "mae_unidades", "N", xCon)
    
    Fg1.SetFocus
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = Fg1.ColIndex("CANTIDAD")
End Sub
 
Sub EliminarItem()
    ' ELIMINA UNA FILA DEL CONTROL FlexGrid Fg1
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
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstIng
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNA SELECCIONADA DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstIng.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If fCierrePeriodo = False Then Exit Sub
        Nuevo
    End If
    If KeyCode = 46 Then
        If fCierrePeriodo = False Then Exit Sub
        Eliminar
    End If
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then VerMovimientos1 IdMenuActivo, NulosN(RstIng("id")), xCon
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case Fg1.ColIndex("CANTIDAD")
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.0000")
    End Select
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Fg1.Editable = flexEDNone: Exit Sub
    
    Select Case Fg1.Col
        Case Fg1.ColIndex("CANTIDAD")
            Fg1.Editable = flexEDKbdMouse
            
        Case Else
            Fg1.Editable = flexEDNone
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case Fg1.ColIndex("CANTIDAD")
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
            
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        AgregarItem
    End If
    If KeyCode = 46 Then
        EliminarItem
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu Menu1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
         Dim NomMes As String
         Dim Cerrado As Boolean
        '------------------------------------------------------------------------------------------
        ' bloqueamos los botones del toolbar
        CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
        '------------------------------------------------------------------------------------------
        pCargarDatos
    End If
End Sub

Private Sub iniciarCampos()
    Dim xRs As New ADODB.Recordset
    
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Mostrando = False
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Transferencia"
    Bloquea
    Blanquea
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = Fg1.FixedRows
    Fg1.SelectionMode = flexSelectionFree
    
    xHorIni = Time
    FechaText.Valor = Date
    FechaText.SetFocus
End Sub

Private Sub Form_Load()
    QueHace = 3
    iniciarCampos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    AgregarItem
End Sub

Private Sub Menu1_3_Click()
    EliminarItem
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstIng.State = 0 Then Exit Sub
        If RstIng.RecordCount = 0 And QueHace <> 1 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstIng.Requery
            Dg1.Refresh
            Cancelar
            
            If RstIng.RecordCount <> 0 Then
                RstIng.MoveFirst
                RstIng.Find "idtransferencia=" & mIdRegistro
                If RstIng.EOF = True Then RstIng.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        RstIng.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 11 Then
        mMesActivo = SeleccionaMes(xCon)
        pCargarDatos
    End If
        
    If Button.Index = 13 Then pExportar
    
    If Button.Index = 16 Then
        Unload Me
        Set RstIng = Nothing
    End If
End Sub

Private Sub pExportar()
    TabOne1.CurrTab = 0
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(4, 3) As String
    
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Fch. Mov.":            xCampos(0, 1) = "fchdoc":           xCampos(0, 2) = 0:  xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Nº Documento":         xCampos(1, 1) = "numdoc":           xCampos(1, 2) = 0:  xCampos(1, 3) = "1200"
    xCampos(2, 0) = "Responsable":          xCampos(2, 1) = "responsable":      xCampos(2, 2) = 0:  xCampos(2, 3) = "2500"
    xCampos(3, 0) = "Alm.Origen":           xCampos(3, 1) = "almorigen":        xCampos(3, 2) = 0:  xCampos(3, 3) = "2500"
    xCampos(4, 0) = "Alm.Destino":          xCampos(4, 1) = "almdestino":       xCampos(4, 2) = 0:  xCampos(4, 3) = "2500"
    '**********************************************************************************************************************************
        
    Set RstTmp = RstIng
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Transferencias", "Periodo: " & LblPeriodo.Caption & "  -  " & AnoTra, "", "Listado de Transferencias", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

Function validarDatos() As Boolean
    If FechaText.Valor = "" Then
        MsgBox "No ha especificado la fecha de movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        FechaText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If Year(FechaText.Valor) <> AnoTra Then
        MsgBox "El año ingresado en la " & Label3(3).Caption & " no coincide con el Ejercicio" & vbCr & "Corrija la fecha o registre en su año que corresponde", vbInformation, xTitulo
        FechaText.Valor = ""
        FechaText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(IdAlmOrigenText.Text) = "" Then
        MsgBox "No ha especificado el Almacen de origen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmOrigenText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(IdAlmDestinoText.Text) = "" Then
        MsgBox "No ha especificado el Almacen de destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmDestinoText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumSerText.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumSerText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumDocText.Text) = "" Then
        MsgBox "No ha especificado el numero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumDocText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumSerOrigenText.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento de origen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumSerOrigenText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumDocOrigenText.Text) = "" Then
        MsgBox "No ha especificado el numero de documento origen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumDocOrigenText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumSerDestinoText.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento de destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumSerDestinoText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumDocDestinoText.Text) = "" Then
        MsgBox "No ha especificado el numero de documento de destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumDocDestinoText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(IdResponsableText.Text) = "" Then
        MsgBox "No ha especificado el responsable el movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdResponsableText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If Fg1.Rows = Fg1.FixedRows Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        validarDatos = False
        Exit Function
    End If
    
    validarDatos = True
End Function
'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Alm_ingreso, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim Transferencia As New AlmacenEntidad.ETransferencia
    Dim A As Integer

On Error GoTo BloqueError
    If Not validarDatos Then Grabar = False: Exit Function
    
    ' Se llenan cabecera
    If QueHace = 1 Then Transferencia.IdTransferencia = 0 Else Transferencia.IdTransferencia = NulosN(RstIng("idtransferencia"))
    Transferencia.FechaTransferencia = Format(FechaText.Valor, "dd/mm/yyyy")
    Transferencia.NumeroSerie = NulosC(NumSerText.Text)
    Transferencia.NumeroDocumento = NulosC(NumDocText.Text)
    Transferencia.IdResponsable = NulosN(IdResponsableText.Text)
    Transferencia.IdEstado = 1
    Transferencia.NumeroSerieOrigen = NulosC(NumSerOrigenText.Text)
    Transferencia.NumeroSerieDestino = NulosC(NumSerDestinoText.Text)
    Transferencia.NumeroDocumentoOrigen = NulosC(NumDocOrigenText.Text)
    Transferencia.NumeroDocumentoDestino = NulosC(NumDocDestinoText.Text)
    Transferencia.IdAlmacenOrigen = NulosN(IdAlmOrigenText.Text)
    Transferencia.IdAlmacenDestino = NulosN(IdAlmDestinoText.Text)
    Transferencia.Glosa = NulosC(GlosaText.Text)
    Transferencia.MesTrabajo = mMesActivo
    Transferencia.AnhoTrabajo = AnoTra
    ' Se llena detalle
    For A = 1 To Fg1.Rows - 1
        Dim TransDet As New AlmacenEntidad.ETransferenciaDet
        
        TransDet.IdTransferenciaDet = Transferencia.IdTransferencia
        TransDet.IdTransferencia = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDTRANSDET")))
        TransDet.IdItem = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDITEM")))
        TransDet.Cantidad = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("CANTIDAD")))
        TransDet.IdUnidadMedida = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDUNIMED")))
        
        Transferencia.TransferenciaDetS.Add TransDet
        Set TransDet = Nothing
    Next A
    
    Set Transferencia.Conexion = xCon
    Transferencia.Save 0, ""

    MsgBox "La trasferencia se registró con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    mIdRegistro = Transferencia.IdTransferencia
    Set Transferencia = Nothing
    Grabar = True
    Exit Function

BloqueError:
    MsgBox "No se pudo registrar transferencia por el siguiente motivo :" + Trim(Err.Description)
    Set Transferencia = Nothing
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA alm_ingreso
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim Transf As New AlmacenEntidad.ETransferencia
    
On Error GoTo BloqueError
    TabOne1.CurrTab = 0
    If RstIng.State = 0 Then Exit Sub
    If RstIng.RecordCount = 0 Then
        MsgBox "No hay Registros de Ingreso/Salida de Almacén para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar el ingreso Nº " + Trim(RstIng("numdoc")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Transf.IdTransferencia = NulosN(RstIng("idtransferencia"))
        Set Transf.Conexion = xCon
        Transf.Delete 0, ""
        RstIng.Requery
        Dg1.Refresh
        MsgBox "La Transferencia se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    Set Transf = Nothing
    Exit Sub
    
BloqueError:
    MsgBox "No se pudo eliminar la transferencia por el siguiente motivo :" + Trim(Err.Description)
    Set Transf = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Transferencia"
    QueHace = 2
    Bloquea
    Blanquea
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = Fg1.FixedRows
    Fg1.Rows = Fg1.Rows + 1
    Fg1.SelectionMode = flexSelectionFree
    MuestraSegundoTab
    xHorIni = Time
    FechaText.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LOS CONTROLES TextBox, PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    Dim A As Integer
    
    FechaText.Valor = ""
    NumSerText.Text = ""
    NumDocText.Text = ""
    IdResponsableText.Text = ""
    ResponsableLabel.Caption = ""
    GlosaText.Text = ""
    IdAlmOrigenText.Text = ""
    AlmOrigenLabel.Caption = ""
    IdAlmDestinoText.Text = ""
    AlmDestinoLabel.Caption = ""
    NumSerOrigenText.Text = ""
    NumDocOrigenText.Text = ""
    NumSerDestinoText.Text = ""
    NumDocDestinoText.Text = ""
    Fg1.Rows = Fg1.FixedRows
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS CONTROLES TEXTBOX, PREPARA PARA AGREGAR O MODIFICAR UN
'*                    REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    FechaText.Locked = Not FechaText.Locked
    NumSerText.Locked = Not NumSerText.Locked
    NumDocText.Locked = Not NumDocText.Locked
    IdResponsableText.Locked = Not IdResponsableText.Locked
    GlosaText.Locked = Not GlosaText.Locked
    IdAlmOrigenText.Locked = Not IdAlmOrigenText.Locked
    IdAlmDestinoText.Locked = Not IdAlmDestinoText.Locked
    NumSerOrigenText.Locked = Not NumSerOrigenText.Locked
    NumDocOrigenText.Locked = Not NumDocOrigenText.Locked
    NumSerDestinoText.Locked = Not NumSerDestinoText.Locked
    NumDocDestinoText.Locked = Not NumDocDestinoText.Locked
    
    habilitar Cmd, Not Cmd(0).Enabled
End Sub

Private Sub IdResponsableText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub IdResponsableText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECLA F5
        Cmd(0).Value = True
    End If
End Sub

Private Sub IdResponsableText_Validate(Cancel As Boolean)
    Dim xRs As New ADODB.Recordset
    
    If NulosC(IdResponsableText.Text) = "" Then Exit Sub
    xRs.CursorLocation = adUseClient
    
    cSQL = "SELECT pla_empleados.id, pla_empleados.nombre AS descripcion " _
        + vbCr + "FROM pla_empleados " _
        + vbCr + "WHERE id = " & NulosN(IdResponsableText.Text) & ""
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        IdResponsableText.Text = ""
        ResponsableLabel.Caption = ""
    Else
        ResponsableLabel.Caption = xRs("descripcion")
    End If
    
    Set xRs = Nothing
End Sub

Private Sub IdAlmOrigenText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub IdAlmOrigenText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Cmd(1).Value = True
    End If
End Sub

Private Sub IdAlmOrigenText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(IdAlmOrigenText.Text) = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    xRs.CursorLocation = adUseClient
    
    cSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion " _
        + vbCr + "FROM alm_almacenes " _
        + vbCr + "WHERE id = " & NulosN(IdAlmOrigenText.Text) & ""
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        IdAlmOrigenText.Text = ""
        AlmOrigenLabel.Caption = ""
    Else
        AlmOrigenLabel.Caption = NulosC(xRs("descripcion"))
    End If
    
    Set xRs = Nothing
End Sub

Private Sub IdAlmDestinoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub IdAlmDestinoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Cmd(2).Value = True
    End If
End Sub

Private Sub IdAlmDestinoText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(IdAlmDestinoText.Text) = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    xRs.CursorLocation = adUseClient
    
    cSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion " _
        + vbCr + "FROM alm_almacenes " _
        + vbCr + "WHERE id = " & NulosN(IdAlmDestinoText.Text) & ""
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        IdAlmDestinoText.Text = ""
        AlmDestinoLabel.Caption = ""
    Else
        AlmDestinoLabel.Caption = NulosC(xRs("descripcion"))
    End If
    
    Set xRs = Nothing
End Sub

'*************************
' NUMEROS DE DOCUMENTO
'*************************
Private Sub NumSerText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumSerText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(NumSerText.Text) <> "" Then
        NumSerText.Text = Format(NumSerText.Text, "0000")
        NumDocText.Text = hallarNumDoc("alm_transferencia", "'" & NulosC(NumSerText.Text) & "'", "numser")
        If NulosC(NumDocText.Text) = "" Then NumSerText.Text = ""
    End If
End Sub

Private Sub NumDocText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumDocText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(NumDocText.Text) = "" Then Exit Sub
    
    NumDocText.Text = Format(NumDocText.Text, "0000000000")
    
    If existeNumeroDoc("alm_transferencia", "'" & NumDocText.Text & "'", "numdoc", "'" & NumSerText.Text & "'", "numser") Then
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        NumDocText.Text = ""
        NumDocText.SetFocus
        Exit Sub
    End If
End Sub

Private Sub NumSerOrigenText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumSerOrigenText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(NumSerOrigenText.Text) <> "" Then
        NumSerOrigenText.Text = Format(NumSerOrigenText.Text, "0000")
        NumDocOrigenText.Text = hallarNumDoc("alm_ingreso", "'" & NulosC(NumSerOrigenText.Text) & "'", "numser")
        If NulosC(NumDocOrigenText.Text) = "" Then NumSerOrigenText.Text = "": Exit Sub
        If NulosC(NumDocDestinoText.Text) <> "" Then
            If (NulosC(NumSerOrigenText.Text) = NulosC(NumSerDestinoText.Text)) And (NulosC(NumDocOrigenText.Text) = NulosC(NumDocDestinoText.Text)) Then
                NumDocOrigenText.Text = NulosN(NumDocDestinoText.Text) + 1
                NumDocOrigenText.Text = Format(NumDocOrigenText.Text, "0000000000")
            End If
        End If
    End If
End Sub

Private Sub NumDocOrigenText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumDocOrigenText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(NumDocText.Text) = "" Then Exit Sub
    
    NumDocOrigenText.Text = Format(NumDocOrigenText.Text, "0000000000")
    
    If existeNumeroDoc("alm_ingreso", "'" & NumDocOrigenText.Text & "'", "numdoc", "'" & NumSerOrigenText.Text & "'", "numser") Then
        If QueHace = 2 Then
            If (NumDocOrigenText.Text = NulosC(RstIng("numdocorig"))) Then Exit Sub
        End If
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        NumDocOrigenText.Text = ""
        NumDocOrigenText.SetFocus
        Exit Sub
    End If
End Sub


Private Sub NumSerDestinoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumSerDestinoText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(NumSerDestinoText.Text) <> "" Then
        NumSerDestinoText.Text = Format(NumSerDestinoText.Text, "0000")
        NumDocDestinoText.Text = hallarNumDoc("alm_ingreso", "'" & NulosC(NumSerDestinoText.Text) & "'", "numser")
        If NulosC(NumDocDestinoText.Text) = "" Then NumSerDestinoText.Text = ""
        If NulosC(NumDocOrigenText.Text) <> "" Then
            If (NulosC(NumSerOrigenText.Text) = NulosC(NumSerDestinoText.Text)) And (NulosC(NumDocOrigenText.Text) = NulosC(NumDocDestinoText.Text)) Then
                NumDocDestinoText.Text = NulosN(NumDocDestinoText.Text) + 1
                NumDocDestinoText.Text = Format(NumDocDestinoText.Text, "0000000000")
            End If
        End If
    End If
End Sub

Private Sub NumDocDestinoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumDocDestinoText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(NumDocText.Text) = "" Then Exit Sub
    
    NumDocDestinoText.Text = Format(NumDocDestinoText.Text, "0000000000")
    
    If existeNumeroDoc("alm_ingreso", "'" & NumDocDestinoText.Text & "'", "numdoc", "'" & NumSerDestinoText.Text & "'", "numser") Then
        If QueHace = 2 Then
            If (NumDocDestinoText.Text = NulosC(RstIng("numdocdest"))) Then Exit Sub
        End If
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        NumDocDestinoText.Text = ""
        NumDocDestinoText.SetFocus
        Exit Sub
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    '********************************************************************
    ' Modificado: 02/04/2012 - Jose Chacon - Modificar referencias a lote
    '********************************************************************
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    If RstIng.RecordCount = 0 Then Exit Sub
    If RstIng.BOF = True Or RstIng.EOF = True Then Exit Sub
    
    cSQL = "SELECT alm_transferencia.* " _
        + vbCr + "FROM alm_transferencia " _
        + vbCr + "WHERE alm_transferencia.idtransferencia=" & NulosN(RstIng("idtransferencia"))
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    FechaText.Valor = xRs("fchdoc")
    NumSerText.Text = NulosC(xRs("numser"))
    NumDocText.Text = NulosC(xRs("numdoc"))
    IdResponsableText.Text = NulosN(xRs("idresponsable"))
    ResponsableLabel.Caption = UCase(Busca_Codigo(NulosN(xRs("idresponsable")), "id", "nombre", "pla_empleados", "N", xCon))
    GlosaText.Text = NulosC(xRs("glosa"))
    IdAlmOrigenText.Text = NulosN(xRs("idalmorig"))
    AlmOrigenLabel.Caption = UCase(Busca_Codigo(NulosN(xRs("idalmorig")), "id", "descripcion", "alm_almacenes", "N", xCon))
    IdAlmDestinoText.Text = NulosN(xRs("idalmdest"))
    AlmDestinoLabel.Caption = UCase(Busca_Codigo(NulosN(xRs("idalmdest")), "id", "descripcion", "alm_almacenes", "N", xCon))
    NumSerOrigenText.Text = NulosC(xRs("numserorig"))
    NumDocOrigenText.Text = NulosC(xRs("numdocorig"))
    NumSerDestinoText.Text = NulosC(xRs("numserdest"))
    NumDocDestinoText.Text = NulosC(xRs("numdocdest"))
        
    cSQL = "SELECT alm_transferenciadet.idtransferenciadet, alm_transferenciadet.idtransferencia, alm_transferenciadet.iditem, alm_transferenciadet.idunimed, alm_transferenciadet.cantidad, alm_transferenciadet.preuni, alm_transferenciadet.lote, alm_transferenciadet.hora, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "FROM (alm_transferenciadet INNER JOIN alm_inventario ON alm_transferenciadet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_transferenciadet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((alm_transferenciadet.idtransferencia) = " & NulosN(RstIng("idtransferencia")) & "));"
    
    Set RstDet = Nothing
    RST_Busq RstDet, cSQL, xCon

    Fg1.Rows = Fg1.FixedRows
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, Fg1.ColIndex("CODIGO")) = NulosC(RstDet("codpro"))
            Fg1.TextMatrix(A, Fg1.ColIndex("ITEM")) = NulosC(RstDet("descripcion"))
            Fg1.TextMatrix(A, Fg1.ColIndex("UM")) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(A, Fg1.ColIndex("STOCK")) = Format(SaldoActual(NulosN(RstDet("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
            Fg1.TextMatrix(A, Fg1.ColIndex("CANTIDAD")) = Format(NulosN(RstDet("cantidad")), FORMAT_CANTIDAD)
            Fg1.TextMatrix(A, Fg1.ColIndex("IDTRANSDET")) = NulosN(RstDet("idtransferenciadet"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDITEM")) = NulosN(RstDet("iditem"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDUNIMED")) = NulosN(RstDet("idunimed"))
            
            RstDet.MoveNext
            
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Sub pCargarDatos()
    TDB_FiltroLimpiar Dg1
    Set RstIng = Nothing
    
    cSQL = "SELECT alm_transferencia.idtransferencia, alm_transferencia.fchdoc, alm_transferencia.idtipdoc, alm_transferencia.numser, alm_transferencia.numdoc, alm_transferencia.idresponsable, alm_transferencia.numserorig, alm_transferencia.numdocorig, alm_transferencia.idalmorig, alm_transferencia.numserdest, alm_transferencia.numdocdest, alm_transferencia.idalmdest, alm_transferencia.idalmdest, alm_transferencia.glosa, alm_almacenes.descripcion AS almorigen, alm_almacenes_1.descripcion AS almdestino, pla_empleados.nombre AS responsable " _
        + vbCr + "FROM ((alm_transferencia INNER JOIN alm_almacenes ON alm_transferencia.idalmorig = alm_almacenes.id) INNER JOIN alm_almacenes AS alm_almacenes_1 ON alm_transferencia.idalmdest = alm_almacenes_1.id) INNER JOIN pla_empleados ON alm_transferencia.idresponsable = pla_empleados.id " _
        + vbCr + "WHERE (((alm_transferencia.ano) = " & AnoTra & ") And ((alm_transferencia.idmes) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY alm_transferencia.fchdoc DESC;"

    RST_Busq RstIng, cSQL, xCon
    Set Dg1.DataSource = RstIng
    
    '********************************************************************************************
    LblPeriodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '********************************************************************************************
End Sub

