VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManTraItem 
   Caption         =   "FrmManTraItem"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TO1 
      Height          =   4395
      Left            =   0
      TabIndex        =   7
      Top             =   345
      Width           =   9555
      _cx             =   16854
      _cy             =   7752
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
      FrontTabForeColor=   -2147483630
      Caption         =   "    &Consulta    |     &Detalle     "
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4020
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   330
         Width           =   9465
         _cx             =   16695
         _cy             =   7091
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
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   4
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmManTraItem.frx":0000
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   2205
            Left            =   300
            TabIndex        =   12
            Top             =   1080
            Width           =   8865
            Begin VB.CommandButton CmdBusCli 
               Height          =   240
               Left            =   2265
               Picture         =   "FrmManTraItem.frx":0070
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   285
               Width           =   240
            End
            Begin VB.TextBox TxtVin4 
               Height          =   285
               Left            =   1335
               Locked          =   -1  'True
               TabIndex        =   5
               Text            =   "TxtVin4"
               Top             =   1650
               Width           =   5055
            End
            Begin VB.TextBox TxtVin3 
               Height          =   285
               Left            =   1335
               Locked          =   -1  'True
               TabIndex        =   4
               Text            =   "TxtVin3"
               Top             =   1365
               Width           =   5055
            End
            Begin VB.TextBox TxtVin2 
               Height          =   285
               Left            =   1335
               Locked          =   -1  'True
               TabIndex        =   3
               Text            =   "TxtVin2"
               Top             =   1080
               Width           =   5055
            End
            Begin VB.TextBox TxtVin1 
               Height          =   285
               Left            =   1335
               Locked          =   -1  'True
               TabIndex        =   2
               Text            =   "TxtVin1"
               Top             =   810
               Width           =   5055
            End
            Begin VB.TextBox TxtIdSeven 
               Height          =   285
               Left            =   1335
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   1
               Text            =   "TxtidSeven"
               Top             =   255
               Width           =   1200
            End
            Begin VB.TextBox TxtId 
               Height          =   285
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   0
               Text            =   "TxtId"
               Top             =   1500
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label LblIdSeven 
               AutoSize        =   -1  'True
               Caption         =   "LblIdSeven"
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
               Height          =   195
               Left            =   6855
               TabIndex        =   25
               Top             =   1170
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label LblDescCue 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDescCue"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2565
               TabIndex        =   24
               Top             =   540
               Width           =   6120
            End
            Begin VB.Label LblNumCue 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNumCue"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1335
               TabIndex        =   23
               Top             =   540
               Width           =   1200
            End
            Begin VB.Label Label4 
               Caption         =   "Cuenta Cont."
               Height          =   210
               Left            =   135
               TabIndex        =   22
               Top             =   585
               Width           =   1095
            End
            Begin VB.Label LblDescripcion 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDescripcion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2565
               TabIndex        =   21
               Top             =   255
               Width           =   6120
            End
            Begin VB.Label Label9 
               Caption         =   "Item Seven"
               Height          =   210
               Left            =   135
               TabIndex        =   19
               Top             =   285
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Vínculo 1"
               Height          =   210
               Left            =   135
               TabIndex        =   18
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label7 
               Caption         =   "Vínculo 2"
               Height          =   210
               Left            =   135
               TabIndex        =   17
               Top             =   1110
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "Vínculo 3"
               Height          =   210
               Left            =   135
               TabIndex        =   16
               Top             =   1395
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Vínculo 4"
               Height          =   210
               Left            =   135
               TabIndex        =   15
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Id"
               Height          =   210
               Left            =   6660
               TabIndex        =   14
               Top             =   1470
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Detalle"
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
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   90
            Width           =   9285
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   4020
         Left            =   -10110
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   9465
         _cx             =   16695
         _cy             =   7091
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
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   2
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmManTraItem.frx":01A2
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   3540
            Left            =   90
            TabIndex        =   10
            Top             =   390
            Width           =   9285
            _ExtentX        =   16378
            _ExtentY        =   6244
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
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codpro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descitem"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Vínculo 1"
            Columns(3).DataField=   "vinc1"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Vínculo 2"
            Columns(4).DataField=   "vinc2"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Vínculo 3"
            Columns(5).DataField=   "vinc3"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Vínculo 4"
            Columns(6).DataField=   "vinc4"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1799"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1720"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7779"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7699"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4921"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4842"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=4763"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4683"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=4948"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=4868"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=5054"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=4974"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Items"
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
            Height          =   240
            Left            =   90
            TabIndex        =   9
            Top             =   90
            Width           =   9285
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":01E4
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":0728
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":0ABA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":0C3E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":1092
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":11AA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":16EE
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":1C32
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":1D46
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":1E5A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":22AE
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":241A
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":2962
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTraItem.frx":2C7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Documento"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar a Excel"
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
   End
End
Attribute VB_Name = "FrmManTraItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim xRstFrm As New ADODB.Recordset
Dim fOrdenLista As Boolean
Dim mIdRegistro& '--identificador del registro
Dim xHorIni As Date
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub CmdBusCli_Click()
    If QueHace = 3 Then Exit Sub
    
    ' EJECUTA LA BUSQUEDA DE UN CLIENTE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Código":        xCampos(0, 1) = "codpro":        xCampos(0, 2) = "1200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripción":   xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "6200":    xCampos(1, 3) = "C"
   '-mostrar solo registros activos y con cuenta contable de compras
    xform.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, con_planctas.cuenta, con_planctas.descripcion AS desccuen " _
        & " FROM con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta Where alm_inventario.activo=-1 and (((alm_inventario.idcuenta) <> 0)) " _
        & " ORDER BY alm_inventario.descripcion"
    
    xform.Titulo = "Buscando Items del seven"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "codpro"
    xform.CampoBusca = "codpro"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblDescripcion.Caption = NulosC(xRs("descripcion"))
        TxtidSeven.Text = NulosC(xRs("codpro"))
        LblIdSeven.Caption = xRs("id")
        
        LblNumCue.Caption = NulosC(xRs("cuenta"))
        LblDescCue.Caption = NulosC(xRs("desccuen"))
        TxtVin1.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TO1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, xRstFrm
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    xRstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
    
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TO1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(xRstFrm("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        'rst_busq xRstFrm, "SELECT tra_item.*, alm_inventario.descripcion FROM tra_item LEFT JOIN alm_inventario ON tra_item.iditem = alm_inventario.id", xCon
        
        RST_Busq xRstFrm, "SELECT tra_item.*,  alm_inventario.codpro, alm_inventario.descripcion AS descitem, con_planctas.cuenta, con_planctas.descripcion AS desccuen " _
            & " FROM con_planctas RIGHT JOIN (tra_item LEFT JOIN alm_inventario ON tra_item.iditem = alm_inventario.id) ON con_planctas.id = alm_inventario.idcuenta", xCon

        'RST_Busq xRstFrm, "SELECT tra_item.*, alm_inventario.descripcion AS descitem, alm_inventario.codpro FROM tra_item LEFT JOIN alm_inventario ON tra_item.iditem = alm_inventario.id;", xCon
        
        Set Dg1.DataSource = xRstFrm
    End If
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Form_Load()
    SetearForm
    SeEjecuto = False
    QueHace = 3
    TO1.CurrTab = 0
End Sub

Sub SetearForm()
    '---------------
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
    TO1.Left = 0
    TO1.Top = 345
    
    TO1.Width = Me.Width - 120
    TO1.Height = Me.Height - 750
    
    Frame1.BackColor = &H8000000F
    Me.Caption = "Transferencia - Mantenimiento de Items"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width <= 3000 Then
        Me.Width = 3000
        Exit Sub
    End If
    
    If Me.Height <= 3000 Then
        Me.Height = 3000
        Exit Sub
    End If
    
    TO1.Width = Me.Width - 120
    TO1.Height = Me.Height - 750
End Sub

Sub Blanquea()
    TxtId.Text = ""
    TxtidSeven.Text = ""
    TxtVin1.Text = ""
    TxtVin2.Text = ""
    TxtVin3.Text = ""
    TxtVin4.Text = ""
    LblDescripcion.Caption = ""
    LblNumCue.Caption = ""
    LblDescCue.Caption = ""
    LblIdSeven.Caption = ""
End Sub

Sub Bloquea()
    'TxtId.Locked = Not TxtId.Locked
    TxtidSeven.Locked = Not TxtidSeven.Locked
    TxtVin1.Locked = Not TxtVin1.Locked
    TxtVin2.Locked = Not TxtVin2.Locked
    TxtVin3.Locked = Not TxtVin3.Locked
    TxtVin4.Locked = Not TxtVin4.Locked
End Sub

Sub MuestraSegundoTab()
    Blanquea
    If xRstFrm.EOF = True Or xRstFrm.BOF = True Or xRstFrm.RecordCount = 0 Then Exit Sub
    
    TxtId.Text = xRstFrm("id")
    If NulosN(xRstFrm("iditem")) = 0 Then
        TxtidSeven.Text = ""
        LblDescripcion.Caption = ""
    Else
        TxtidSeven.Text = NulosC(xRstFrm("codpro"))
        LblIdSeven.Caption = xRstFrm("iditem")
        LblDescripcion.Caption = NulosC(xRstFrm("descitem"))
    End If
    TxtVin1.Text = NulosC(xRstFrm("vinc1"))
    TxtVin2.Text = NulosC(xRstFrm("vinc2"))
    TxtVin3.Text = NulosC(xRstFrm("vinc3"))
    TxtVin4.Text = NulosC(xRstFrm("vinc4"))
    
    LblNumCue.Caption = NulosC(xRstFrm("cuenta"))
    LblDescCue.Caption = NulosC(xRstFrm("desccuen"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TO1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Sub Nuevo()
    xHorIni = Time
    QueHace = 1
    Label2.Caption = "Agregando Item"
    TO1.CurrTab = 1
    TO1.TabEnabled(0) = False
    ActivaTool
    Blanquea
    Bloquea
    TxtId.Text = HallaCodigoTabla("tra_item", xCon, "id")
    
    TxtidSeven.SetFocus
End Sub

Sub Modificar()
    xHorIni = Time
    QueHace = 2
    Label2.Caption = "Modificando Item"
    TO1.CurrTab = 1
    TO1.TabEnabled(0) = False
    ActivaTool
    Blanquea
    Bloquea
    MuestraSegundoTab
    TxtidSeven.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    If xRstFrm.State = 0 Then Exit Sub
    If xRstFrm.EOF = True Or xRstFrm.BOF = True Or xRstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿ Esta seguro de eliminar el registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM tra_item WHERE id = " & xRstFrm("id") & ""
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xRstFrm("id") & " AND idform = " & IdMenuActivo
        
        xRstFrm.Requery
        Dg1.Refresh
        MsgBox "El registro fue eliminado con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Function Grabar() As Boolean
    Dim xId As Double
    Dim xRst As New ADODB.Recordset
    Dim xSql As String
    
    If QueHace = 1 Then
        xSql = "SELECT tra_item.vinc1, tra_item.vinc2, tra_item.vinc3, tra_item.vinc4, tra_item.id, tra_item.iditem " _
            & " From tra_item " _
            & " WHERE (((tra_item.vinc1)='" & NulosC(TxtVin1.Text) & "') AND ((tra_item.vinc2)='" & NulosC(TxtVin2.Text) & "') " _
            & " AND ((tra_item.vinc3)='" & NulosC(TxtVin3.Text) & "') AND ((tra_item.vinc4)='" & NulosC(TxtVin4.Text) & "'));"
    
        RST_Busq xRst, xSql, xCon
        
        If xRst.RecordCount <> 0 Then
            MsgBox "Ya existe un registro con la informacion ingresada, el Registro es el Nº " & Format(xRst("id"), "0000"), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xRst = Nothing
            Exit Function
        End If
        
        Set xRst = Nothing
    End If
    
    
    If NulosN(LblIdSeven.Caption) = 0 Then
        MsgBox "Falta especifica el Item del seven", vbInformation, xTitulo
        TxtidSeven.SetFocus
    End If
    
On Error GoTo LaCague
    
    Dim xCampos(5, 5) As String
    
    xCon.BeginTrans
    
    'ESPECIFICAMOS EL ID DEL MOVIMIENTO
    If QueHace = 1 Then
        xId = HallaCodigoTabla("tra_item", xCon, "id")
    Else
        xId = xRstFrm("id")
    End If
    
    mIdRegistro = xId

    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
    '--------------------------------
    'GRABAMOS LA CABECERA DE LA LETRA
    xCampos(0, 0) = "id":        xCampos(0, 1) = Str(xId):                   xCampos(0, 2) = "S":    xCampos(0, 3) = "N":     xCampos(0, 4) = "": xCampos(0, 5) = "S"
    xCampos(1, 0) = "iditem":    xCampos(1, 1) = NulosN(LblIdSeven.Caption): xCampos(1, 2) = "N":     xCampos(1, 3) = "N":     xCampos(1, 4) = "Falta especificar el Item de seven"
    xCampos(2, 0) = "vinc1":     xCampos(2, 1) = TxtVin1.Text:               xCampos(2, 2) = "N":    xCampos(2, 3) = "C":     xCampos(2, 4) = ""
    xCampos(3, 0) = "vinc2":     xCampos(3, 1) = TxtVin2.Text:               xCampos(3, 2) = "N":    xCampos(3, 3) = "C":     xCampos(3, 4) = ""
    xCampos(4, 0) = "vinc3":     xCampos(4, 1) = TxtVin3.Text:               xCampos(4, 2) = "N":    xCampos(4, 3) = "C":     xCampos(4, 4) = ""
    xCampos(5, 0) = "vinc4":     xCampos(5, 1) = TxtVin4.Text:               xCampos(5, 2) = "N":    xCampos(5, 3) = "C":     xCampos(5, 4) = ""
    
    If QueHace = 1 Then
        If EscribirNuevoRegistro(xCampos, "tra_item", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Else
        If ModificarRegistro(xCampos, "tra_item", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    End If
    
    '---------------------------------------------------------------------------
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
        
    MsgBox "El registro se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Grabar = True
    
    xCon.CommitTrans
    Grabar = True
    Exit Function
    
LaCague:
Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Err.Clear
    Grabar = False
End Function

Sub Cancelar()
    Label2.Caption = "Detalle"
    Bloquea
    TO1.TabEnabled(0) = True
    TO1.CurrTab = 0
    QueHace = 3
    ActivaTool
    Dg1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            xRstFrm.Requery
            Dg1.Refresh
            
            '--Posiocionar en registro actual
            If xRstFrm.RecordCount <> 0 Then xRstFrm.MoveFirst
            xRstFrm.Find "id = " & mIdRegistro & ""
            If xRstFrm.EOF = True Then xRstFrm.MoveFirst
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TDB_Actualizar Me, TO1, Dg1, xRstFrm
'        xRstFrm.Filter = adFilterNone
'        TDB_FiltroLimpiar Dg1
        
        
    End If
'
    If Button.Index = 13 Then
        Dim xFun As New eps_librerias.FuncionesDGrid
        xFun.xNomEmp = "" 'NomEmp
        xFun.xNumRuc = "" 'NumRUC
        xFun.ExportarDGExcel xRstFrm, Dg1, "LISTA DE ITEMS PARA IMPORTACION"
        Set xFun = Nothing
    End If

    If Button.Index = 16 Then
        Set xRstFrm = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtidSeven_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtidSeven_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtidSeven.Text) = "" Then
        LblIdSeven.Caption = 0
        TxtidSeven.Text = ""
        LblDescripcion.Caption = ""
        Exit Sub
    End If
    
    Dim xRs As New ADODB.Recordset
    Dim xSql As String
    
    xSql = "SELECT alm_inventario.id, alm_inventario.descripcion, con_planctas.cuenta, con_planctas.descripcion AS desccuen" _
        & " FROM con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta WHERE alm_inventario.codpro = '" & NulosC(TxtidSeven.Text) & "' ORDER BY alm_inventario.descripcion"

    RST_Busq xRs, xSql, xCon
    If xRs.RecordCount <> 0 Then
        LblDescripcion.Caption = NulosC(xRs("descripcion"))
        LblIdSeven.Caption = NulosN(xRs("id"))
    Else
        TxtidSeven.Text = ""
        LblDescripcion.Caption = ""
        LblIdSeven.Caption = 0
    End If
    Set xRs = Nothing
End Sub

Private Sub TxtVin1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtVin2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtVin3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtVin4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
