VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMantEmpresa1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEVEN - Mantenimiento de Empresa"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   Icon            =   "FrmMantEmpresa1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9930
      Top             =   1710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
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
            Picture         =   "FrmMantEmpresa1.frx":030A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":0BE0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":0D64
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":11B8
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":12D0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":1814
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":1D58
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":1E6C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":1F80
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":23D4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantEmpresa1.frx":2540
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5220
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   9870
      _cx             =   17410
      _cy             =   9208
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4800
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   9780
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4425
            Left            =   60
            TabIndex        =   26
            Top             =   330
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   7805
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº R.U.C."
            Columns(1).DataField=   "numruc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Empresa"
            Columns(2).DataField=   "nomemp"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Ruta B.D."
            Columns(3).DataField=   "ruta"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Estado"
            Columns(4).DataField=   "estado"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=979"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=6297"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6218"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4577"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4498"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1535"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1455"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consultando Empresas"
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
            Index           =   0
            Left            =   75
            TabIndex        =   15
            Top             =   30
            Width           =   9645
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   9810
            TabIndex        =   14
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4800
         Left            =   10515
         TabIndex        =   11
         Top             =   375
         Width           =   9780
         Begin VB.Frame Frame3 
            Height          =   3705
            Left            =   210
            TabIndex        =   17
            Top             =   660
            Width           =   9390
            Begin VB.TextBox TxtRaiz2 
               Height          =   300
               Left            =   3735
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   36
               Text            =   "TxtRaiz2"
               Top             =   2070
               Width           =   1095
            End
            Begin VB.TextBox TxtRaiz1 
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   34
               Text            =   "TxtRaiz1"
               Top             =   2070
               Width           =   1095
            End
            Begin VB.Frame FraRuta 
               Caption         =   "[ Data Origen ]"
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
               Height          =   1080
               Left            =   3600
               TabIndex        =   27
               Top             =   210
               Visible         =   0   'False
               Width           =   5700
               Begin VB.CommandButton CmdBusOrigen 
                  Height          =   240
                  Left            =   1515
                  Picture         =   "FrmMantEmpresa1.frx":2A88
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  Top             =   330
                  Width           =   240
               End
               Begin VB.TextBox TxtIdOrigen 
                  Height          =   300
                  Left            =   855
                  Locked          =   -1  'True
                  TabIndex        =   29
                  Text            =   "TxtIdOrigen"
                  Top             =   300
                  Width           =   930
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Ruta"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   33
                  Top             =   705
                  Width           =   345
               End
               Begin VB.Label LblRuta 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblRuta"
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
                  Left            =   855
                  TabIndex        =   32
                  Top             =   630
                  Width           =   4545
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Empresa"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   31
                  Top             =   375
                  Width           =   615
               End
               Begin VB.Label LblOrigen 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblOrigen"
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
                  Left            =   1800
                  TabIndex        =   30
                  Top             =   300
                  Width           =   3585
               End
            End
            Begin VB.Frame fra 
               Caption         =   "( Sincronizar )"
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
               Height          =   780
               Index           =   1
               Left            =   4560
               TabIndex        =   25
               Top             =   2805
               Visible         =   0   'False
               Width           =   3450
               Begin VB.OptionButton Opt_Sincroniza 
                  Caption         =   "Si"
                  Height          =   195
                  Index           =   0
                  Left            =   660
                  TabIndex        =   8
                  Top             =   345
                  Width           =   765
               End
               Begin VB.OptionButton Opt_Sincroniza 
                  Caption         =   "No"
                  Height          =   195
                  Index           =   1
                  Left            =   2070
                  TabIndex        =   9
                  Top             =   345
                  Width           =   765
               End
            End
            Begin VB.Frame fra 
               Caption         =   "( Activo )"
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
               Height          =   780
               Index           =   0
               Left            =   375
               TabIndex        =   24
               Top             =   2805
               Width           =   3270
               Begin VB.OptionButton Opt_Activo 
                  Caption         =   "Si"
                  Height          =   195
                  Index           =   0
                  Left            =   600
                  TabIndex        =   6
                  Top             =   345
                  Width           =   765
               End
               Begin VB.OptionButton Opt_Activo 
                  Caption         =   "No"
                  Height          =   195
                  Index           =   1
                  Left            =   1980
                  TabIndex        =   7
                  Top             =   345
                  Width           =   765
               End
            End
            Begin VB.TextBox TxtAño 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   1
               Text            =   "TxtA"
               Top             =   675
               Width           =   885
            End
            Begin VB.TextBox TxtRutaBD 
               BackColor       =   &H8000000F&
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   5
               Text            =   "TxtRutaBD"
               Top             =   2415
               Width           =   7035
            End
            Begin VB.TextBox TxtNumRuc 
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   11
               TabIndex        =   2
               Text            =   "TxtNumRuc"
               Top             =   1020
               Width           =   1440
            End
            Begin VB.TextBox TxtNomCorto 
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   4
               Text            =   "TxtNomCorto"
               Top             =   1710
               Width           =   4500
            End
            Begin VB.TextBox TxtNomEmp 
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   3
               Text            =   "TxtNomEmp"
               Top             =   1365
               Width           =   4500
            End
            Begin VB.TextBox TxtCodigo 
               Height          =   300
               Left            =   2010
               Locked          =   -1  'True
               MaxLength       =   13
               TabIndex        =   0
               Text            =   "TxtCodigo"
               Top             =   315
               Width           =   885
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Raiz2"
               Height          =   195
               Index           =   3
               Left            =   3270
               TabIndex        =   37
               Top             =   2175
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Raiz1"
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   35
               Top             =   2175
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Año de Trabajo"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   23
               Top             =   780
               Width           =   1095
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Empresa"
               Height          =   195
               Index           =   7
               Left            =   360
               TabIndex        =   22
               Top             =   1470
               Width           =   1215
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Corto"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   21
               Top             =   1815
               Width           =   975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº R.U.C."
               Height          =   195
               Index           =   6
               Left            =   360
               TabIndex        =   20
               Top             =   1125
               Width           =   705
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Ruta"
               Height          =   195
               Index           =   8
               Left            =   360
               TabIndex        =   19
               Top             =   2520
               Width           =   345
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Código"
               Height          =   195
               Index           =   10
               Left            =   360
               TabIndex        =   18
               Top             =   420
               Width           =   495
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Empresa"
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
            Left            =   60
            TabIndex        =   12
            Top             =   30
            Width           =   9675
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmMantEmpresa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMMANTEMPRESA
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO PARA EL MANTENIMIENTO DE EMPRESAS
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 03/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstEmp As New ADODB.Recordset          ' RECORSER PRINCIPAL
Dim QueHace As Integer                     ' VARIABLE PARA IDENTIFICAR LAS ACCIONES SOBRE EL FORMULARIO (1 = NUEVO,2 = MODIFICAR, 3 = SOLOLECTURA)
Dim SeEjecuto As Integer                   ' VARIABLE QUE INDICARA SI EL FORMULARIO YA EJECUTO EL EVENTO LOAD
Dim xId As Double                         ' VARIABLE QUE ALMACENARA EL ID DE LOS REGISTRO
Dim xConRuta As ADODB.Connection
Dim fOrdenLista As Boolean                 ' --especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


'*****************************************************************************************************
'* Nombre Modulo  : MuestraSegundoTab()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : MUESTRA LOS DATOS AL DETALLE DE LA EMPRESA SELECCIONADA
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub MuestraSegundoTab()

    Blanquea
    If RstEmp.EOF = True Or RstEmp.BOF = True Or RstEmp.RecordCount = 0 Then Exit Sub
    
    TxtAño.Text = NulosC(RstEmp("anotra"))
    TxtCodigo.Text = RstEmp("id")
    TxtNumRuc.Text = NulosC(RstEmp("numruc"))
    TxtNomEmp.Text = NulosC(RstEmp("nomemp"))
    TxtNomCorto.Text = NulosC(RstEmp("abrevia"))
    
    Dim xRuta As String
    xRuta = NulosC(RstEmp("ruta"))
    TxtRaiz1.Text = Mid(xRuta, 1, InStr(1, xRuta, "\") - 1)
    xRuta = Mid(xRuta, InStr(1, xRuta, "\") + 1)
    
    TxtRaiz2.Text = Mid(xRuta, 1, InStr(1, xRuta, "\") - 1)
    xRuta = Mid(xRuta, InStr(1, xRuta, "\") + 1)
    
    TxtRutaBD.Text = AP_RUTABD & NulosC(RstEmp("ruta"))
    
    If NulosN(RstEmp("activo")) = -1 Then
        Opt_Activo(0).Value = True
        Opt_Activo(1).Value = False
    Else
        Opt_Activo(0).Value = False
        Opt_Activo(1).Value = True
    End If
    
    If NulosN(RstEmp("sincronizar")) = -1 Then
        Opt_Sincroniza(0).Value = True
        Opt_Sincroniza(1).Value = False
    Else
        Opt_Sincroniza(0).Value = False
        Opt_Sincroniza(1).Value = True
    End If
    
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstEmp
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstEmp.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstEmp("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO QUE SE EJECUTAR AL CARGA EL FORMULARIO, AQUI SE CARGARAN EN EL RECORSET PRINCIPAL LAS
    ' EMPRESAS REGISTRADAS Y SERAN MOSTRADAS EN EL DATAGRID DEL FORMULARIO
    
    If SeEjecuto = False Then
        Dim Rpta As Integer
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = 86
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        '--oculatar boton eliminar
        Toolbar1.Buttons(3).Visible = False
                
        ' ABRIMOS LA CONECCION A LA BD DE ENLACE PARA PODER REALIZARLAS OPERACIONES
        Dim xFun As New eps_librerias.FuncionesData
        
        xFun.F_BASEDATOS = AP_RUTABD + "data.mdb"                                           ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
        xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
        xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
        xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
        xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
        
        Set xConRuta = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
        Set xFun = Nothing
        
        ' CARGAMOS LOS DATOS DE LA EMPRESA EN EL RECORSET
        RST_Busq RstEmp, "SELECT IIf([activo]=-1,'Activo','Inactivo') AS estado, IIf([sincronizar]=-1,'Si','No') AS sincroniza,* From mae_empresa ORDER BY mae_empresa.id", xConRuta

        Set Dg1.DataSource = RstEmp
        
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : ActivaToolbar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub ActivaToolbar()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Nuevo()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : LE DICE AL FORMULARIO QUE SE AGREGARA UN NUEVO REGISTRO, PARA ELLO ACTUALIZA EL
'*                  EL VALOR DE LA VARIABLE QUEHACE = 1
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Nuevo()
    Dim xAño As String
    QueHace = 1
    xHorIni = Time
    Label5.Caption = "Agregando Empresa"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaToolbar
    Blanquea
    Bloquea True
    TxtCodigo.SetFocus
    
    xId = HallaCodigoTabla("mae_empresa", xConRuta, "id")
    TxtCodigo.Text = xId
    xAño = Format(Date, "YYYY")
    TxtAño.Text = xAño
    
    TxtRaiz1.Text = xAño
    TxtRaiz2.Text = Format(xId, "0000")
    
    TxtRutaBD.Text = AP_RUTABD & TxtRaiz1.Text + "\" + TxtRaiz2.Text + "\data.mdb"
    
    FraRuta.Visible = True
    
    Opt_Activo(0).Value = True
    Opt_Sincroniza(1).Value = True
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Nuevo()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : LE DICE AL FORMULARIO QUE SE MODIFICARA UN REGISTRO, PARA ELLO ACTUALIZA EL
'*                  VALOR DE LA VARIABLE QUEHACE = 2
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    xHorIni = Time
    Label5.Caption = "Modificando Empresa"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaToolbar
    Blanquea
    Bloquea True
    TxtCodigo.SetFocus
    MuestraSegundoTab
    xId = RstEmp("id")
    TxtCodigo.SetFocus
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTARA AL CARGAR EL EVENTO LOAD
    SeEjecuto = False
    QueHace = 3
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
End Sub

Private Sub Opt_Activo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    If Opt_Sincroniza(0).Value = True Then Opt_Sincroniza(0).SetFocus
    If Opt_Sincroniza(1).Value = True Then Opt_Sincroniza(1).SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstEmp.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then
        TDB_Actualizar Me, TabOne1, Dg1, RstEmp
'        TabOne1.CurrTab = 0
'        TDB_FiltroLimpiar Dg1
'        RstEmp.Filter = ""
    End If
    
    If Button.Index = 9 Then
        Set RstEmp = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Grabar()
'* Tipo           : FUNCCION
'* Descripcion    : PERMITE GUARDAR LOS DATOS EDITADOS EN EL FORMULARIO, RETORANA UN VALOR VERDADERO
'*                  CUANDO EL REGISTRO SE GUARDA CON EXITO
'* Paranetros     : NULL
'* Retorna        : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim A As New FileSystemObject
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SON LOS CORRECTOS
    If NulosN(TxtAño.Text) = 0 Then
        MsgBox "No ha especificado el año de trabajo para la empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtAño.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumRuc.Text) = "" Then
        MsgBox "No ha especificado el numero de ruc de la empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If Len(NulosC(TxtNumRuc.Text)) < 11 Then
        MsgBox "Nº de R.U.C. inválidos, ingreso un Nº de R.U.C. valido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNomEmp.Text) = "" Then
        MsgBox "No ha especificado el nombre de la empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNomEmp.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNomCorto.Text) = "" Then
        MsgBox "No ha especificado el nombre corto de la empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNomCorto.SetFocus
        Exit Function
    End If
    '--Verificar si ruta existe
    If QueHace = 1 Then
        If A.FileExists(NulosC(AP_RUTABD) + NulosC(LblRuta.Caption)) = False Then
            MsgBox "La data origen no existe, seleccione otra data origen", vbInformation, xTitulo
            TxtIdOrigen.Text = ""
            LblRuta.Caption = ""
            LblOrigen.Caption = ""
            TxtIdOrigen.SetFocus
            Exit Function
        End If
    End If
    
    Dim RstCab As New ADODB.Recordset
        
    ' PREGUNTAMOS QUE ACCION ESTA EFECTUANDO EL REGISTRO
    If QueHace = 1 Then
        ' SI ES NUEVO REGISTRO VERIFICAMOS QUE LA EMPRESA NO ESTE REGISTRADA
        RST_Busq RstCab, "SELECT * FROM mae_empresa WHERE anotra = " & NulosN(TxtAño.Text) & " AND numruc = '" & NulosC(TxtNumRuc.Text) & "'", xConRuta
        If RstCab.RecordCount <> 0 Then
            ' SI LA EMPRESA FUE REGISTRA EMITIMOS UN AVISO Y SALIMOS DE LA FUNCION
            MsgBox "La empresa ya fue registrada en el año de trabajo " + NulosC(TxtAño.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set RstCab = Nothing
            Exit Function
        End If
        Set RstCab = Nothing
    End If
    
    On Error GoTo lAcAGUE
    xConRuta.BeginTrans
    
    ' GRAMAOS LOS DATOS
    If QueHace = 1 Then
        ' OBTENEMOS EL ID PARA EL NUEVO REGITROS
        xId = HallaCodigoTabla("mae_empresa", xConRuta, "id")
        TxtCodigo.Text = xId
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM mae_empresa", xConRuta
        
        ' AGREGAMOS UN NUEVO REGISTRO
        RstCab.AddNew
        RstCab("id") = xId
    Else
        ' BUSCAMOS EL REGISTRO Y TRAEMOS LOS DATOS
        xId = RstEmp("id")
        RST_Busq RstCab, "SELECT * FROM mae_empresa WHERE id  = " & xId & "", xConRuta
    End If
    
    ' ASIGNAMOS LOS DATOS A CADA CAMPO
    RstCab("numruc") = Trim(TxtNumRuc.Text)
    RstCab("nomemp") = TxtNomEmp.Text
    RstCab("abrevia") = TxtNomCorto.Text
    RstCab("anotra") = NulosN(TxtAño.Text)
    
    If QueHace = 1 Then
        RstCab("ruta") = TxtRaiz1.Text & "\" & TxtRaiz2.Text & "\data.mdb"
    End If
'    If QueHace = 1 Then
'        RstCab("ruta") = NulosC(TxtAño.Text) + "\" + Format(xId, "0000") + "\data.mdb"
'    Else
'        RstCab("ruta") = Trim(TxtRutaBD.Text)
'    End If
    
    If Opt_Activo(0).Value = True Then
        RstCab("activo") = -1
    Else
        RstCab("activo") = 0
    End If
    
    If Opt_Sincroniza(0).Value = True Then
        RstCab("sincronizar") = -1
    Else
        RstCab("sincronizar") = 0
    End If
        
    RstCab.Update
    Set RstCab = Nothing
    
    xConRuta.CommitTrans
    
    If QueHace = 1 Then
        ' SI ES NUEVO CREAMOS LA CARPETA PARA ALOJAR LA BASE DE DATOS
        Dim xRutaEmp As String
        Dim xRutaAño As String
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        Dim xArchivo As String
        Dim RutaMaestro As String
        
        'xRutaEmp = NulosC(AP_RUTABD) + NulosC(TxtAño.Text) + "\" + Format(xId, "0000")
        xRutaEmp = NulosC(AP_RUTABD) & TxtRaiz1.Text & "\" & TxtRaiz2.Text & "\"
               
        'verificamos si existe la carpeta con el año especificado
        If A.FolderExists(NulosC(AP_RUTABD) + NulosC(TxtRaiz1.Text)) = True Then
            'preguntamos si existe la carpeta que vamos a crear
            If A.FolderExists(xRutaEmp) = False Then
                'creamos la carpeta
                A.CreateFolder (xRutaEmp)
            End If
        Else
            'creamo la carpeta del año especificado
            A.CreateFolder (NulosC(AP_RUTABD) + NulosC(TxtRaiz1.Text))
            'preguntamos si existe la carpeta que vamos a crear
            If A.FolderExists(xRutaEmp) = False Then
                'creamos la carpeta
                A.CreateFolder (xRutaEmp)
            End If
        End If
        
'        RST_Busq Rst, "SELECT * FROM mae_empresa WHERE maestro = -1", xConRuta
        RutaMaestro = NulosC(AP_RUTABD) + NulosC(LblRuta.Caption)
        
        xArchivo = NulosC(AP_RUTABD) + NulosC(TxtRaiz1.Text) + "\" + NulosC(TxtRaiz2.Text) + "\data.mdb"
                       
        If A.FileExists(xArchivo) = False Then
            A.CopyFile RutaMaestro, xRutaEmp
        Else
            '--Hacer copia de seguridad
            FileCopy xArchivo, NulosC(AP_RUTABD) + NulosC(TxtRaiz1.Text) + "\" + NulosC(TxtRaiz2.Text) + "\data-" & Time & ".mdb"
            '--copiar contenido
            A.CopyFile RutaMaestro, xRutaEmp
        End If
        '--Creando el archivo de texto con el nombre de la empresa creada
        
        A.CreateTextFile NulosC(AP_RUTABD) + NulosC(TxtRaiz1.Text) + "\" + NulosC(TxtRaiz2.Text) & "\" & TxtNomEmp.Text & " " & TxtAño.Text & ".txt", True
        Set A = Nothing
        
        '--Pendiente Limpiar data
        
        
    End If
    
    Grabar = True
    If QueHace = 1 Then
        MsgBox "La nueva empresa se generó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        '--Limpiar data de nueva empresa
        LimpiarData
    Else
        MsgBox "La empresa se modificó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    Exit Function
    
lAcAGUE:
    ''''Resume
    xConRuta.RollbackTrans
    Set RstCab = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + vbCr + Trim(Err.Description), vbCritical, xTitulo
End Function

'*****************************************************************************************************
'* Nombre Modulo  : Cancelar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : PERMITE CANELAR EL PROCESO DE INGRESO O MODIFICACION DE REGISTRO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    ActivaToolbar
    Bloquea False
    Label5.Caption = "Detalle de la Empresa"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub TxtAño_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtAño_Validate(Cancel As Boolean)
    If QueHace <> 1 Then Exit Sub
    If TxtAño.Text = "" Then
        TxtRaiz1.Text = ""
    Else
        TxtRaiz1.Text = Format(TxtAño.Text, "0000")
        TxtRaiz1_Validate True
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNomCorto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtNomEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRaiz1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRaiz1_Validate(Cancel As Boolean)
    If QueHace <> 1 Then Exit Sub
    If TxtRaiz1.Text = "" Then
        TxtRutaBD.Text = "#ERROR - INGRESE RAIZ1"
    Else
        TxtRutaBD.Text = AP_RUTABD & TxtRaiz1.Text + "\" + TxtRaiz2.Text + "\data.mdb"
    End If

End Sub

Private Sub TxtRaiz2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRaiz2_Validate(Cancel As Boolean)
    If QueHace <> 1 Then Exit Sub
    If TxtRaiz2.Text = "" Then
        TxtRutaBD.Text = "#ERROR - INGRESE RAIZ2"
    Else
        TxtRaiz2.Text = Format(TxtRaiz2.Text, "0000")
        TxtRutaBD.Text = AP_RUTABD & TxtRaiz1.Text + "\" + TxtRaiz2.Text + "\data.mdb"
    End If

End Sub

Private Sub TxtRutaBD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Blanquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : BLANQUEA LOS CONTROLES DEL FORMULARIO PARA EL INGRESO DE DATOS
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Blanquea()
    TxtAño.Text = ""
    TxtCodigo.Text = ""
    TxtNumRuc.Text = ""
    TxtNomEmp.Text = ""
    TxtNomCorto.Text = ""
    TxtRutaBD.Text = ""
    
    '---
    TxtIdOrigen.Text = ""
    LblOrigen.Caption = ""
    LblRuta.Caption = ""
    FraRuta.Visible = False

End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Bloquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Bloquea(xBand As Boolean)
    TxtAño.Locked = Not xBand
    TxtNumRuc.Locked = Not xBand
    TxtNomEmp.Locked = Not xBand
    TxtNomCorto.Locked = Not xBand
    TxtRutaBD.Locked = Not xBand
    habilitar fra, xBand
    TxtRaiz1.Locked = Not xBand
    TxtRaiz2.Locked = Not xBand
    If QueHace = 1 Then FraRuta.Visible = Not xBand
End Sub

Private Sub TxtIdOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdOrigen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusOrigen_Click
    End If
    If KeyCode = 117 Then
        TxtIdOrigen.Text = ""
        LblOrigen.Caption = ""
    End If
End Sub

Private Sub TxtIdOrigen_Validate(Cancel As Boolean)
    '--solo validar cuando sea nuevo registro
    If QueHace <> 1 Then Exit Sub
    If NulosN(TxtIdOrigen.Text) = 0 Then
        LblOrigen.Caption = ""
        LblRuta.Caption = ""
        Exit Sub
    End If
    LblOrigen.Caption = Busca_Codigo(NulosN(TxtIdOrigen.Text), "id", "nomemp", "mae_empresa", "N", xConRuta)
    LblRuta.Caption = Busca_Codigo(NulosN(TxtIdOrigen.Text), "id", "ruta", "mae_empresa", "N", xConRuta)
End Sub


Private Sub CmdBusOrigen_Click()
    If QueHace <> 1 Then Exit Sub
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Cod.":          xCampos(0, 1) = "id":      xCampos(0, 2) = "500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Ruc":           xCampos(1, 1) = "numruc":  xCampos(1, 2) = "1200":  xCampos(1, 3) = "C"
    xCampos(2, 0) = "Razón Social":  xCampos(2, 1) = "nomemp":  xCampos(2, 2) = "3000":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "Año":           xCampos(3, 1) = "anotra":  xCampos(3, 2) = "700":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "ruta":          xCampos(4, 1) = "ruta":    xCampos(4, 2) = "2200":  xCampos(4, 3) = "C"
    
    CARGAR_DLL_EPSBUSCAR xConRuta, xRs, "select * from mae_empresa", xCampos(), "Seleccionar origen de Datos", "ruta desc", "nomemp", CualquierParte
    
    If xRs.State = 1 Then
        If xRs.RecordCount > 0 Then
            TxtIdOrigen.Text = xRs("id")
            LblOrigen.Caption = NulosC(xRs("nomemp"))
            LblRuta.Caption = NulosC(xRs("ruta"))
            TxtIdOrigen_Validate False
        End If
    End If
    Set xRs = Nothing
End Sub

Private Sub LimpiarData()
    Dim xConTmp As New ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    
    xFun.F_BASEDATOS = TxtRutaBD.Text                                             ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
    
    Set xConTmp = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
    Set xFun = Nothing

    If xConTmp.State = 1 Then
        '--Limpiar Ingreso/Salida Almacén
        xConTmp.Execute "Delete From alm_ingresodet;"
        xConTmp.Execute "Delete From alm_ingresodoc;"
        xConTmp.Execute "Delete From alm_ingreso;"
        
        '--Limpiar Compras
        xConTmp.Execute "Delete From com_comprasdet;"
        xConTmp.Execute "Delete From com_comprascosto;"
        xConTmp.Execute "Delete From com_compras;"
        
        '--Limpiar Honorarios
        xConTmp.Execute "Delete From com_honorarioscosto;"
        xConTmp.Execute "Delete From com_honorariosdet;"
        xConTmp.Execute "Delete From com_honorariosuneg;"
        xConTmp.Execute "Delete From com_honorarios;"
        
        '--Limpiar Reembolsables
        xConTmp.Execute "Delete From com_reembolsables;"
        
        '--Limpiar Canjes de Documentos
        xConTmp.Execute "Delete From con_canjesdet;"
        xConTmp.Execute "Delete From con_canjes;"
        
        '--Limpiar Detraccion
        xConTmp.Execute "Delete From con_detraccion;"
        
        '--Limpiar Percepcion
        xConTmp.Execute "Delete From con_percepciondet;"
        xConTmp.Execute "Delete From con_percepcion;"
        
        '--Limpiar Provisiones
        xConTmp.Execute "Delete From con_provicionesdetdoc;"
        xConTmp.Execute "Delete From con_provicionesdet;"
        xConTmp.Execute "Delete From con_proviciones;"
        
        '--Limpiar Retencion
        xConTmp.Execute "Delete From con_retenciondet;"
        xConTmp.Execute "Delete From con_retencion;"
        
        '--Limpiar Letras en Cartera
        xConTmp.Execute "Delete From let_letradet;"
        xConTmp.Execute "Delete From let_letradoc;"
        xConTmp.Execute "Delete From let_letra;"
        
        '--Limpiar Planilla de Letras
        xConTmp.Execute "Delete From let_planilladet;"
        xConTmp.Execute "Delete From let_planilla;"
        
        '--Limpiar Boleta de Pago
        xConTmp.Execute "Delete From pla_boletadet;"
        xConTmp.Execute "Delete From pla_boletacosto;"
        xConTmp.Execute "Delete From pla_boleta;"
        
        '--Limpiar Conciliaciones
        xConTmp.Execute "Delete From tes_concidet;"
        xConTmp.Execute "Delete From tes_conci;"
        
        '--Limpiar Tesoreria Ingresos-Egresos
        xConTmp.Execute "Delete From tes_cajadestinodet;"
        xConTmp.Execute "Delete From tes_cajadestino;"
        xConTmp.Execute "Delete From tes_cajaorigendet;"
        xConTmp.Execute "Delete From tes_cajaori;"
        xConTmp.Execute "Delete From tes_caja;"
        
        '--Limpiar Varios
        xConTmp.Execute "Delete From var_analisisctacte;"
        xConTmp.Execute "Delete From var_edicion;"
        xConTmp.Execute "Delete From var_importados;"
        xConTmp.Execute "Delete From var_ordendespacho;"
        
        '--Limpiar Gasto Debito
        xConTmp.Execute "Delete From vta_gastodebitodet;"
        xConTmp.Execute "Delete From vta_gastodebito;"
        
        '--Limpiar Guia de Remisión
        xConTmp.Execute "Delete From vta_guiadet;"
        xConTmp.Execute "Delete From vta_guia;"
        
        '--Limpiar Pedido
        xConTmp.Execute "Delete From vta_pedidodet;"
        xConTmp.Execute "Delete From vta_pedido;"
        
        '--Limpiar Proforma
        xConTmp.Execute "Delete From vta_proforma;"
        
        '--Limpiar Ventas
        xConTmp.Execute "Delete From vta_ventasdet;"
        xConTmp.Execute "Delete From vta_ventas;"
        
        '--Limpiar Diario
        xConTmp.Execute "Delete From con_diario;"
        
        xConTmp.Execute "UPDATE mae_empresa SET mae_empresa.anotra = " & NulosN(TxtAño.Text) & ", mae_empresa.numruc = '" & NulosC(TxtNumRuc.Text) & "', mae_empresa.nomemp = '" & NulosC(TxtNomEmp.Text) & "',  " _
            & " mae_empresa.diremp = '', mae_empresa.repleg = '', mae_empresa.numdocrepleg = '', mae_empresa.iddocrepleg = 0; "
        
        MsgBox "Proceda a verificar los datos de la nueva empresa en [?\ Setup] ", vbInformation, xTitulo
        
        
    End If
    xConTmp.Close

End Sub
