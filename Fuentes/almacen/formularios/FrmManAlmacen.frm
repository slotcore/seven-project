VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacen - Mantenimiento Almacen"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlmacen.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6600
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   10920
      _cx             =   19262
      _cy             =   11642
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6180
         Left            =   45
         TabIndex        =   4
         Top             =   375
         Width           =   10830
         Begin VB.Frame FrmReceta 
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
            Height          =   5745
            Left            =   60
            TabIndex        =   7
            Top             =   390
            Width           =   10685
            Begin VB.TextBox TxtCodigo 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "TxtCodigo"
               Top             =   1680
               Width           =   3945
            End
            Begin VB.CommandButton cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2745
               Picture         =   "FrmManAlmacen.frx":2B10
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   3555
               Width           =   240
            End
            Begin VB.CommandButton cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   2745
               Picture         =   "FrmManAlmacen.frx":2C42
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   3210
               Width           =   240
            End
            Begin VB.TextBox txtDesc 
               Height          =   735
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "txtDesc"
               Top             =   2370
               Width           =   8250
            End
            Begin VB.TextBox TxtIdAlm 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   8
               Text            =   "TxtIdAlm"
               Top             =   2040
               Width           =   8250
            End
            Begin VB.TextBox TxtTipAlm 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   14
               Text            =   "TxtTipAlm"
               Top             =   3180
               Width           =   1215
            End
            Begin VB.TextBox TxtMetVal 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   18
               Text            =   "TxtMetVal"
               Top             =   3525
               Width           =   1215
            End
            Begin VB.Label lblIdMetVal 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblIdMetVal"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   9720
               TabIndex        =   24
               Top             =   120
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblIdtipalm 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblIdtipalm"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   8880
               TabIndex        =   23
               Top             =   120
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Codigo"
               Height          =   195
               Index           =   2
               Left            =   600
               TabIndex        =   22
               Top             =   1695
               Width           =   495
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Metodo Valoriz."
               Height          =   195
               Index           =   1
               Left            =   540
               TabIndex        =   20
               Top             =   3555
               Width           =   1095
            End
            Begin VB.Label LblMetVal 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblMetVal"
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
               Left            =   3060
               TabIndex        =   19
               Top             =   3540
               Width           =   6990
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Alm."
               Height          =   195
               Index           =   0
               Left            =   540
               TabIndex        =   16
               Top             =   3210
               Width           =   885
            End
            Begin VB.Label LblTipoAlm 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoAlm"
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
               Left            =   3060
               TabIndex        =   15
               Top             =   3180
               Width           =   6990
            End
            Begin VB.Label lblidalm 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidalm"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   8160
               TabIndex        =   12
               Top             =   120
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Glosa"
               Height          =   195
               Index           =   1
               Left            =   600
               TabIndex        =   10
               Top             =   2430
               Width           =   405
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion"
               Height          =   195
               Index           =   0
               Left            =   585
               TabIndex        =   9
               Top             =   2040
               Width           =   840
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Almacen"
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
            Left            =   105
            TabIndex        =   5
            Top             =   30
            Width           =   10655
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6180
         Left            =   -11475
         TabIndex        =   1
         Top             =   375
         Width           =   10830
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5760
            Left            =   30
            TabIndex        =   2
            Top             =   345
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   10160
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
            Columns(1).Caption=   "Codigo"
            Columns(1).DataField=   "codigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo"
            Columns(3).DataField=   "destipalm"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=9763"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=9684"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=3493"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3413"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Almacen"
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
            Left            =   105
            TabIndex        =   3
            Top             =   30
            Width           =   10655
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   609
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Item"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar un Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar Item"
               EndProperty
            EndProperty
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir "
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar             "
      End
   End
End
Attribute VB_Name = "FrmManAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMANALAMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : AQUI SE CREAN, MODIFICAN Y ELIMINAN LOS ITEMS Y SE LES ASIGNA LA CUENTA CONTABLE.
'*                  : Y EL CENTRO DE COSTO.
'* DISEÑADO POR     : Jose Chacon
'* ULTIMA REVISION  : 24/04/2012
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstPro As New ADODB.Recordset
Dim xHorIni As Date
Dim fOrdenLista As Boolean
Dim IdMenuActivo As Integer
Dim mIdRegistro&
Dim cSQL As String
Dim Agregando As Boolean
Dim F As New SistemaLogica.Funciones

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Bloquea()
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    txtDesc.Locked = Not txtDesc.Locked
    TxtCodigo.Locked = Not TxtCodigo.Locked
    TxtTipAlm.Locked = Not TxtTipAlm.Locked
    TxtMetVal.Locked = Not TxtMetVal.Locked
    habilitar cmd, Not cmd(0).Enabled
End Sub

Private Sub iniciarCampos()
    TabOne1.CurrTab = 0
    
    CaracteresNumericos = "0123456789." & Chr(8)
End Sub

Private Sub pCargarGrid()
    cSQL = "SELECT alm_almacenes.id, alm_almacenes.codigo, alm_almacenes.descripcion, alm_almacenes.obs, alm_almacenes.idtipalm, alm_almacenes.idmetval, mae_tipoalmacen.codigo AS codtipalm, mae_tipoalmacen.descripcion AS destipalm, mae_configval.abrev AS codmetval, mae_configval.descripcion AS desmetval " _
        + vbCr + "FROM (alm_almacenes LEFT JOIN mae_tipoalmacen ON alm_almacenes.idtipalm = mae_tipoalmacen.id) LEFT JOIN mae_configval ON alm_almacenes.idmetval = mae_configval.id " _
        + vbCr + "ORDER BY alm_almacenes.descripcion"
    
    RST_Busq RstPro, cSQL, xCon

    Set Dg1.DataSource = RstPro
End Sub

Sub Blanquea()
    lblidalm.Caption = 0
    lblIdtipalm.Caption = 0
    lblIdMetVal.Caption = 0
    TxtCodigo.Text = ""
    TxtIdAlm.Text = ""
    txtDesc.Text = ""
    TxtTipAlm.Text = ""
    TxtMetVal.Text = ""
    LblTipoAlm.Caption = ""
    LblMetVal.Caption = ""
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim nSQLId As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim A As Integer
    
    Select Case Index
        Case 0 ' Tipo de Almacen
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "codigo":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                   
            cSQL = "SELECT mae_tipoalmacen.id, mae_tipoalmacen.codigo, mae_tipoalmacen.descripcion " _
                + vbCr + "FROM mae_tipoalmacen"
                        
            nTitulo = "Buscando Ítems"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblIdtipalm.Caption = NulosN(xRs("id"))
            TxtTipAlm.Text = NulosC(xRs("codigo"))
            LblTipoAlm.Caption = NulosC(xRs("descripcion"))
            TxtMetVal.SetFocus
        
        Case 1 ' Metodo de Valorizacion
            ReDim xCampos(2, 4) As String
                
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "abrev":          xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
                   
            cSQL = "SELECT mae_configval.* " _
                + vbCr + "FROM mae_configval"
                        
            nTitulo = "Buscando Metodo de Valorizacion"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblIdMetVal.Caption = NulosN(xRs("id"))
            TxtMetVal.Text = NulosC(xRs("abrev"))
            LblMetVal.Caption = NulosC(xRs("descripcion"))
    End Select
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstPro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstPro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    ' SI SE HA PRESIONADO LA TECLA F12 MOSTRAMOS LA INFORMACION DE EDICION DEL REGISTRO
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPro("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    ' CARGAMOS LOS ITEMS DEL INVENTARIO Y LOS MOSTRAMOS EN LA LA PRIMERA PESTAÑA DEL FORMULARIO, ESTE EVENTO SOLO SE EJECUTARA
    ' UNA SOLA VEZ
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        SeEjecuto = True
        pCargarGrid
        
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Almacen"
    xHorIni = Time
    TxtIdAlm.SetFocus
End Sub

Private Sub Form_Load()
    ' CARGAMOS EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    iniciarCampos
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 3 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstPro.Requery
            Dg1.Refresh
            Cancelar
            
            If RstPro.RecordCount <> 0 Then
                RstPro.MoveFirst
                RstPro.Find "id=" & mIdRegistro
                If RstPro.EOF = True Then RstPro.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstPro.Filter = ""
    End If
    
    If Button.Index = 15 Then
        Set RstPro = Nothing
        Unload Me
    End If
End Sub

Function validarDatos() As Boolean
    If NulosC(TxtIdAlm.Text) = "" Then
        MsgBox "No ha especificado un nombre para el Almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdAlm.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(TxtCodigo.Text) = "" Then
        MsgBox "No ha especificado el codigo del Almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCodigo.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(TxtTipAlm.Text) = "" Then
        MsgBox "No ha especificado el tipo del Almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipAlm.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    validarDatos = True
End Function

Function Grabar() As Boolean
    Dim Rpta As Integer
    Dim Almacen As New AlmacenEntidad.EAlmacen
        
On Error GoTo LaCague
    If Not validarDatos Then Grabar = False: Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Item", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    
    On Error GoTo LaCague
    If QueHace = 1 Then Almacen.IdAlmacen = 0 Else Almacen.IdAlmacen = NulosN(RstPro("id"))
        
    Almacen.Codigo = NulosC(TxtCodigo.Text)
    Almacen.Descripcion = NulosC(TxtIdAlm.Text)
    Almacen.Glosa = NulosC(txtDesc.Text)
    Almacen.IdTipoAlmacen = NulosN(lblIdtipalm.Caption)
    Almacen.IdMetodoValorizacion = NulosN(lblIdMetVal.Caption)
    
    Set Almacen.Conexion = xCon
    If Not Almacen.Save(0, "") Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el registro"
    
    mIdRegistro = Almacen.IdAlmacen
    MsgBox "El item se grabo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Set Almacen = Nothing
    Exit Function
    
LaCague:
    Grabar = False
    Set Almacen = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
End Function

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELAR EL PROCESO DE AGRGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA UN REGISTRO PARA SU MODIFICACION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Bloquea
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        Blanquea
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If

    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Almacen"
    xHorIni = Time
    TxtIdAlm.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA EL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim Almacen As New AlmacenEntidad.EAlmacen
    
    ' SI EL ITEM NO TIENE NINGUNA OPERACION SE PROCEDE A ELIMINAR PREVIA AUTORIZACION DEL USUARIO
    Rpta = MsgBox("¿ Esta seguro de eliminar el registro ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Almacen.IdAlmacen = NulosN(RstPro("id"))
        Set Almacen.Conexion = xCon
        Almacen.Delete CLng(xIdUsuario), F.MachineName
        
        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPro.Requery
        Dg1.Refresh
        Exit Sub
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DEL REGISTRO ACTUAL, ESTE EVENTO SE EJECUTA CUANDO EL
'*                    FORMULARIO ESTA EN MODO DE LECTURA O MODIFICAR
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    lblidalm.Caption = NulosN(RstPro("id"))
    lblIdtipalm = NulosN(RstPro("idtipalm"))
    lblIdMetVal = NulosN(RstPro("idmetval"))
    TxtCodigo.Text = NulosC(RstPro("codigo"))
    TxtIdAlm.Text = NulosC(RstPro("descripcion"))
    txtDesc.Text = NulosC(RstPro("obs"))
    TxtTipAlm.Text = NulosC(RstPro("codtipalm"))
    LblTipoAlm.Caption = NulosC(RstPro("destipalm"))
    TxtMetVal.Text = NulosC(RstPro("codmetval"))
    LblMetVal.Caption = NulosC(RstPro("desmetval"))
End Sub
