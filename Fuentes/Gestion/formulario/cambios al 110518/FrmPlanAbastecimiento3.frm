VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanAbastecimiento3 
   Caption         =   "Compras - Plan de Abastecimiento"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11310
   ScaleWidth      =   14610
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   8730
      Left            =   165
      TabIndex        =   1
      Top             =   480
      Width           =   12465
      _cx             =   21987
      _cy             =   15399
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
      FrontTabForeColor=   -2147483630
      Caption         =   "   &Consulta   |   &Detalle   "
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
      Begin SizerOneLibCtl.ElasticOne Eo1 
         Height          =   8310
         Left            =   -13020
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   375
         Width           =   12375
         _cx             =   21828
         _cy             =   14658
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
         BackColor       =   12640511
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
         _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0000
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   7770
            Left            =   90
            TabIndex        =   3
            Top             =   450
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   13705
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Proyecto"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripcion"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Ini"
            Columns(2).DataField=   "fchini"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Fin"
            Columns(3).DataField=   "fchfin"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Estado"
            Columns(4).DataField=   "estado"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8202"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8123"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1799"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1720"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1667"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1588"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H400000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
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
            Caption         =   "Plan de Abastecimiento"
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
            Height          =   300
            Left            =   90
            TabIndex        =   4
            Top             =   90
            Width           =   12195
         End
      End
      Begin SizerOneLibCtl.ElasticOne Eo10 
         Height          =   8310
         Left            =   45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   375
         Width           =   12375
         _cx             =   21828
         _cy             =   14658
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
         BackColor       =   12648447
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
         GridRows        =   3
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0043
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   7035
            Left            =   90
            TabIndex        =   20
            Top             =   1185
            Width           =   12195
            _cx             =   21511
            _cy             =   12409
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
            FrontTabForeColor=   -2147483630
            Caption         =   "   &Terminados   |   &Intermedios  "
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
            Begin SizerOneLibCtl.ElasticOne Eo11 
               Height          =   6615
               Left            =   45
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   45
               Width           =   12105
               _cx             =   21352
               _cy             =   11668
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
               BackColor       =   12648384
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
               GridRows        =   3
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0091
               Begin VSFlex7Ctl.VSFlexGrid Fg4 
                  Height          =   3075
                  Left            =   90
                  TabIndex        =   28
                  Top             =   3450
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5424
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
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":00E2
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
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   3030
                  Left            =   90
                  TabIndex        =   27
                  Top             =   90
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5345
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
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":01DA
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
               Begin SizerOneLibCtl.ElasticOne Eo12 
                  Height          =   210
                  Left            =   90
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   3180
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   370
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
                  GridRows        =   1
                  GridCols        =   2
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0274
                  Begin VB.CommandButton Command4 
                     Height          =   30
                     Left            =   5955
                     Picture         =   "FrmPlanAbastecimiento3.frx":02B4
                     Style           =   1  'Graphical
                     TabIndex        =   31
                     Top             =   90
                     Width           =   5880
                  End
                  Begin VB.CommandButton Command2 
                     Height          =   30
                     Left            =   90
                     Picture         =   "FrmPlanAbastecimiento3.frx":03F2
                     Style           =   1  'Graphical
                     TabIndex        =   30
                     Top             =   90
                     Width           =   5805
                  End
               End
            End
            Begin SizerOneLibCtl.ElasticOne Eo14 
               Height          =   6615
               Left            =   -12750
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   45
               Width           =   12105
               _cx             =   21352
               _cy             =   11668
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
               BackColor       =   16761087
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
               GridRows        =   3
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0530
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   3060
                  Left            =   90
                  TabIndex        =   26
                  Top             =   3465
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5397
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
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":0581
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
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   3045
                  Left            =   90
                  TabIndex        =   25
                  Top             =   90
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5371
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
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":0679
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
               Begin SizerOneLibCtl.ElasticOne Eo15 
                  Height          =   210
                  Left            =   90
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   3195
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   370
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
                  GridRows        =   1
                  GridCols        =   2
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0773
                  Begin VB.CommandButton Command3 
                     Height          =   30
                     Left            =   5955
                     Picture         =   "FrmPlanAbastecimiento3.frx":07B3
                     Style           =   1  'Graphical
                     TabIndex        =   32
                     Top             =   90
                     Width           =   5880
                  End
                  Begin VB.CommandButton Command1 
                     Height          =   30
                     Left            =   90
                     Picture         =   "FrmPlanAbastecimiento3.frx":08F1
                     Style           =   1  'Graphical
                     TabIndex        =   29
                     Top             =   90
                     Width           =   5805
                  End
               End
            End
         End
         Begin SizerOneLibCtl.ElasticOne Eo13 
            Height          =   735
            Left            =   90
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   390
            Width           =   12195
            _cx             =   21511
            _cy             =   1296
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
            GridRows        =   1
            GridCols        =   3
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0A2F
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   555
               Left            =   7185
               TabIndex        =   14
               Top             =   90
               Width           =   4920
               Begin VB.CommandButton CmdAdd 
                  Caption         =   "Agregar Plan de Produccion"
                  Height          =   525
                  Left            =   1275
                  TabIndex        =   16
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.CommandButton CmdVerEst 
                  Caption         =   "&Ver Estacionalidad"
                  Height          =   525
                  Left            =   75
                  TabIndex        =   15
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Registros : "
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   3900
                  TabIndex        =   18
                  Top             =   165
                  Width           =   1020
               End
               Begin VB.Label LblNumReg 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "LblNumReg"
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
                  Height          =   210
                  Left            =   3900
                  TabIndex        =   17
                  Top             =   375
                  Width           =   1020
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   555
               Left            =   90
               TabIndex        =   7
               Top             =   90
               Width           =   6135
               Begin VB.TextBox TxtDesc 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   945
                  Locked          =   -1  'True
                  TabIndex        =   8
                  Text            =   "TxtDesc"
                  Top             =   75
                  Width           =   5385
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
                  Height          =   300
                  Left            =   945
                  TabIndex        =   9
                  Top             =   390
                  Width           =   1305
                  _ExtentX        =   2302
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
                  Valor           =   "06/02/2006"
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
                  Height          =   300
                  Left            =   4650
                  TabIndex        =   10
                  Top             =   390
                  Width           =   1305
                  _ExtentX        =   2302
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
                  Valor           =   "06/02/2006"
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Termino"
                  Height          =   195
                  Left            =   3465
                  TabIndex        =   13
                  Top             =   420
                  Width           =   930
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripcion"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   12
                  Top             =   105
                  Width           =   840
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Inicio"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   11
                  Top             =   420
                  Width           =   735
               End
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Plan de Produccion"
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
            TabIndex        =   19
            Top             =   90
            Width           =   12195
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
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
            Picture         =   "FrmPlanAbastecimiento3.frx":0A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":0FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":1146
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":16B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":213A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":224E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":2362
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":27B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":2922
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
      Width           =   14610
      _ExtentX        =   25770
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar plan de abastecimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar plan de abastecimiento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar plan de abastecimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar plan de abastecimiento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Programa de Produccion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lista total de insumos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmPlanAbastecimiento3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstInsumos As New ADODB.Recordset
Dim RstPlanAbas As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim xHorIni As Date                 'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xMesInicio As Integer

Private Sub CmdAdd_Click()
    ' PERMITE AGREGAR UN PLAN DE PRODUCCION AL PLAN DE ABASTECIMIENTO QUE SE CREANDO O EDITANDO
        
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    Dim cSQL As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"

    cSQL = "SELECT ges_plaprod.id, ges_plaprod.descripcion, ges_plaprod.fchini, ges_plaprod.fchfin, ges_plaprod.mesini, ges_plaprod.año " _
        + vbCr + "From ges_plaprod " _
        + vbCr + "ORDER BY ges_plaprod.descripcion;"
    
    xform.SQLCad = cSQL
    
    xform.Titulo = "Buscando Plan de Produccion"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Dim xId As Integer
        
        xId = xRs("id")
        TxtFchIni.Valor = xRs("fchini")
        TxtFchFin.Valor = xRs("fchfin")
        MostrarTerminados xId, xRs("mesini"), xRs("año")
        MostrarIntermedios xId, xRs("mesini"), xRs("año")
        TabOne2.CurrTab = 0
        PintarGrid
        Set xform = Nothing
        Set xRs = Nothing
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub MostrarTerminados(IdPlanProduccion As Integer, MesIni As Integer, AñoTra As Integer)
    Dim xSQL As String
    Dim xMes As String
    Dim xRst As New ADODB.Recordset
    Dim A, B, xMesIni, xAñoTra As Integer
    Dim xStock, xDiferencia As Double
    
    xMesIni = MesIni
    xAñoTra = AñoTra
    
    ' MOSTRAMOS LOS PRODUCTOS TERMINADOS
    xSQL = "TRANSFORM Sum(ges_plaproddet.cantidad) AS SumaDecantidad " _
        & " SELECT ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, " _
        & " Sum(ges_plaproddet.cantidad) AS [total] " _
        & " FROM ges_plaproddet LEFT JOIN (alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON ges_plaproddet.codpro = alm_inventario.id " _
        & " Where (((ges_plaproddet.idpv) = " & IdPlanProduccion & ")) " _
        & " GROUP BY ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet.idmes"
    
    RST_Busq xRst, xSQL, xCon
    
    Fg1.Cols = 5
    Fg2.Cols = 5
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg1.Cols = Fg1.Cols + 1
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg1.Cols = Fg1.Cols + 1
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Programado"
        Fg1.ColWidth(Fg1.Cols - 1) = 1100
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Stock"
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Diferencia"
        Fg1.ColWidth(Fg1.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = xRst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRst("abrev")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg1.TextMatrix(0, B), 1, 2))))
                Fg1.TextMatrix(Fg1.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
            
            xStock = SaldoActual(NulosN(Fg1.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 3) = Format(xRst("total"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 2) = Format(xStock, "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 1) = Format(xRst("total") - xStock, "0.00")
    
            xRst.MoveNext
        Next A
    End If
   
    
    ' MOSTRAMOS LOS INSUMOS DE LOS TEMINADOS
    xSQL = "TRANSFORM Sum(terminados.ins_totins) AS SumaDeins_totins " _
        & " SELECT terminados.ins_iditem, terminados.ins_desc, terminados.ins_idunimed, terminados.ins_unimed, Sum(terminados.ins_totins) AS [total] " _
        & " FROM " _
        & " ( " _
        & "     SELECT ges_plaproddet.idpv, ges_plaproddet.codpro, ges_plaproddet.idmes, ges_plaproddet.cantidad AS pro_can, alm_inventario.descripcion AS pro_desc, " _
        & "     pro_recetains.iditem AS ins_iditem, alm_inventario_1.descripcion AS ins_desc, pro_recetains.idunimed AS ins_idunimed, mae_unidades.abrev AS ins_unimed, " _
        & "     pro_recetains.canpro, [pro_recetains].[canpro]*[ges_plaproddet].[cantidad] AS ins_totins " _
        & "     FROM (ges_plaproddet LEFT JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) LEFT JOIN (((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) " _
        & "     LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_recetains.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades ON alm_inventario_1.idunimed = mae_unidades.id) " _
        & "     ON ges_plaproddet.codpro = pro_receta.iditem " _
        & "     Where (((ges_plaproddet.idpv) =  " & IdPlanProduccion & ") And ((pro_receta.prirec) = 1) And ((alm_inventario_1.tippro) = 1 Or (alm_inventario_1.tippro) = 4)) " _
        & " ) AS terminados " _
        & " GROUP BY terminados.ins_iditem, terminados.ins_desc, terminados.ins_idunimed, terminados.ins_unimed PIVOT terminados.idmes"

    
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg2.Cols = Fg2.Cols + 1
            Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg2.Cols = Fg2.Cols + 1
            Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Programado"
        Fg2.ColWidth(Fg2.Cols - 1) = 1100
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Stock"
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Diferencia"
        Fg2.ColWidth(Fg2.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRst("ins_iditem")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRst("ins_idunimed")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = xRst("ins_desc")
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = xRst("ins_unimed")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))))
                Fg2.TextMatrix(Fg2.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
                                   
            xStock = SaldoActual(NulosN(Fg2.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 3) = Format(xRst("total"), "0.00")
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 2) = Format(xStock, "0.00")
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 1) = Format(xRst("total") - xStock, "0.00")
            
            xRst.MoveNext
        Next A
    End If
    
End Sub



Sub MostrarIntermedios(IdPlanProduccion As Integer, MesIni As Integer, AñoTra As Integer)
    Dim xSQL As String
    Dim xMes As String
    Dim xRst As New ADODB.Recordset
    Dim A, B, xMesIni, xAñoTra As Integer
    Dim xStock, xDiferencia As Double
    
    xMesIni = MesIni
    xAñoTra = AñoTra
    
    ' MOSTRAMOS LOS PRODUCTOS TERMINADOS
'    xSql = "TRANSFORM Sum(ges_plaproddet.cantidad) AS SumaDecantidad " _
'        & " SELECT ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, " _
'        & " Sum(ges_plaproddet.cantidad) AS [total] " _
'        & " FROM ges_plaproddet LEFT JOIN (alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON ges_plaproddet.codpro = alm_inventario.id " _
'        & " Where (((ges_plaproddet.idpv) = " & IdPlanProduccion & ")) " _
'        & " GROUP BY ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
'        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet.idmes"
    xSQL = "TRANSFORM Sum(ges_plaproddet2.cantidad) AS SumaDecantidad " _
        & " SELECT ges_plaproddet2.idpv, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, ges_plaproddet2.codpro, Sum(ges_plaproddet2.cantidad) AS [total] " _
        & " FROM ges_plaproddet2 LEFT JOIN (alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON ges_plaproddet2.codpro = alm_inventario.id " _
        & " Where (((ges_plaproddet2.idpv) = " & IdPlanProduccion & ")) " _
        & " GROUP BY ges_plaproddet2.idpv, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, ges_plaproddet2.codpro" _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet2.idmes"
    
    RST_Busq xRst, xSQL, xCon
    
    Fg3.Cols = 5
    Fg3.Rows = 1
    Fg4.Cols = 5
    Fg4.Rows = 1
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg3.Cols = Fg3.Cols + 1
            Fg3.TextMatrix(0, Fg3.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg3.Cols = Fg3.Cols + 1
            Fg3.TextMatrix(0, Fg3.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Programado"
        Fg3.ColWidth(Fg3.Cols - 1) = 1100
        
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Stock"
        
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Diferencia"
        Fg3.ColWidth(Fg3.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = xRst("codpro")
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = xRst("idunimed")
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = xRst("descripcion")
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = xRst("abrev")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg3.TextMatrix(0, B), 1, 2))))
                Fg3.TextMatrix(Fg3.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
            
            xStock = SaldoActual(NulosN(Fg3.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 3) = Format(xRst("total"), "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 2) = Format(xStock, "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 1) = Format(xRst("total") - xStock, "0.00")
    
            xRst.MoveNext
        Next A
    End If
   
    
    ' MOSTRAMOS LOS INSUMOS DE LOS TEMINADOS
'    xSql = "TRANSFORM Sum(terminados.ins_totins) AS SumaDeins_totins " _
'        & " SELECT terminados.ins_iditem, terminados.ins_desc, terminados.ins_idunimed, terminados.ins_unimed, Sum(terminados.ins_totins) AS [total] " _
'        & " FROM " _
'        & " ( " _
'        & "     SELECT ges_plaproddet.idpv, ges_plaproddet.codpro, ges_plaproddet.idmes, ges_plaproddet.cantidad AS pro_can, alm_inventario.descripcion AS pro_desc, " _
'        & "     pro_recetains.iditem AS ins_iditem, alm_inventario_1.descripcion AS ins_desc, pro_recetains.idunimed AS ins_idunimed, mae_unidades.abrev AS ins_unimed, " _
'        & "     pro_recetains.canpro, [pro_recetains].[canpro]*[ges_plaproddet].[cantidad] AS ins_totins " _
'        & "     FROM (ges_plaproddet LEFT JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) LEFT JOIN (((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) " _
'        & "     LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_recetains.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades ON alm_inventario_1.idunimed = mae_unidades.id) " _
'        & "     ON ges_plaproddet.codpro = pro_receta.iditem " _
'        & "     Where (((ges_plaproddet.idpv) =  " & IdPlanProduccion & ") And ((pro_receta.prirec) = 1) And ((alm_inventario_1.tippro) = 1 Or (alm_inventario_1.tippro) = 4)) " _
'        & " ) AS terminados " _
'        & " GROUP BY terminados.ins_iditem, terminados.ins_desc, terminados.ins_idunimed, terminados.ins_unimed PIVOT terminados.idmes"

    xSQL = "TRANSFORM Sum(intermedio.ins_totins) AS SumaDeins_totins " _
        & " SELECT intermedio.ins_iditem, intermedio.ins_desc, intermedio.ins_idunimed, intermedio.ins_unimed, Sum(intermedio.ins_totins) AS [total] " _
        & " FROM " _
        & " ( " _
        & "     SELECT ges_plaproddet2.idpv, ges_plaproddet2.codpro, ges_plaproddet2.idmes, ges_plaproddet2.cantidad AS pro_can, alm_inventario.descripcion AS pro_desc, " _
        & "     pro_recetains.iditem AS ins_iditem, alm_inventario_1.descripcion AS ins_desc, alm_inventario_1.idunimed AS ins_idunimed, mae_unidades.abrev AS ins_unimed, " _
        & "     pro_recetains.canpro, [pro_recetains].[canpro]*[ges_plaproddet2].[cantidad] AS ins_totins " _
        & "     FROM ((((ges_plaproddet2 LEFT JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) LEFT JOIN pro_receta ON ges_plaproddet2.codpro = pro_receta.iditem) " _
        & "     LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_recetains.iditem = alm_inventario_1.id) " _
        & "     LEFT JOIN mae_unidades ON alm_inventario_1.idunimed = mae_unidades.id " _
        & "     Where (((ges_plaproddet2.idpv) = " & IdPlanProduccion & ") And ((pro_receta.prirec) = 1) And ((alm_inventario_1.tippro) = 1 Or (alm_inventario_1.tippro) = 4)) " _
        & " ) AS intermedio " _
        & " GROUP BY intermedio.ins_iditem, intermedio.ins_desc, intermedio.ins_idunimed, intermedio.ins_unimed ORDER BY intermedio.ins_desc PIVOT intermedio.idmes"
   
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg4.Cols = Fg4.Cols + 1
            Fg4.TextMatrix(0, Fg4.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg4.Cols = Fg4.Cols + 1
            Fg4.TextMatrix(0, Fg4.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Programado"
        Fg4.ColWidth(Fg4.Cols - 1) = 1100
        
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Stock"
        
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Diferencia"
        Fg4.ColWidth(Fg4.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(Fg4.Rows - 1, 1) = xRst("ins_iditem")
            Fg4.TextMatrix(Fg4.Rows - 1, 2) = xRst("ins_idunimed")
            Fg4.TextMatrix(Fg4.Rows - 1, 3) = xRst("ins_desc")
            Fg4.TextMatrix(Fg4.Rows - 1, 4) = xRst("ins_unimed")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg4.TextMatrix(0, B), 1, 2))))
                Fg4.TextMatrix(Fg4.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
                                   
            xStock = SaldoActual(NulosN(Fg4.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 3) = Format(xRst("total"), "0.00")
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 2) = Format(xStock, "0.00")
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 1) = Format(xRst("total") - xStock, "0.00")
            
            xRst.MoveNext
        Next A
    End If
    
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Form_Activate()
'Modificado: 08/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios

    ' SEGUNDO EVENTO A EJECUTARSE DESPUES DE CARGARSE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        
        RST_Busq RstPlanAbas, "SELECT IIf([ges_planaba]![activo]=0,'No Activo','Activo') AS estado, * " _
            & " From ges_planaba ORDER BY ges_planaba.id desc", xCon
        
        Set Dg1.DataSource = RstPlanAbas

    End If
End Sub


Sub PintarGrid()
    Dim A As Integer
    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 3, Fg1.Rows - 1, Fg1.Cols - 3, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 2, Fg1.Rows - 1, Fg1.Cols - 2, &HC0FFC0, flexFillRepeat

    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 3, Fg2.Rows - 1, Fg2.Cols - 3, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 2, Fg2.Rows - 1, Fg2.Cols - 2, &HC0FFC0, flexFillRepeat
    
    GRID_COLOR_FONDO Fg3, 1, Fg3.Cols - 3, Fg3.Rows - 1, Fg3.Cols - 3, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg3, 1, Fg3.Cols - 2, Fg3.Rows - 1, Fg3.Cols - 2, &HC0FFC0, flexFillRepeat

    GRID_COLOR_FONDO Fg4, 1, Fg4.Cols - 3, Fg4.Rows - 1, Fg4.Cols - 3, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg4, 1, Fg4.Cols - 2, Fg4.Rows - 1, Fg4.Cols - 2, &HC0FFC0, flexFillRepeat
        
    ' ALINEAMOS LOS ENCABEZADOS DELAS COLUMNAS
    For A = 1 To Fg1.Cols - 1
        Fg1.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg1, 0, A, , True, &H8000000F, Fg1.TextMatrix(0, A)
    Next A
    
    For A = 1 To Fg2.Cols - 1
        Fg2.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg2, 0, A, , True, &H8000000F, Fg2.TextMatrix(0, A)
    Next A
    
    For A = 1 To Fg3.Cols - 1
        Fg3.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg3, 0, A, , True, &H8000000F, Fg3.TextMatrix(0, A)
    Next A
    
    For A = 1 To Fg4.Cols - 1
        Fg4.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg4, 0, A, , True, &H8000000F, Fg4.TextMatrix(0, A)
    Next A

    
    Fg1.FrozenCols = 4
    Fg2.FrozenCols = 4
    Fg3.FrozenCols = 4
    Fg4.FrozenCols = 4
    
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
    'Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Private Sub Modificar()
    QueHace = 2
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label1.Caption = "Modificando Plan de Abastecimiento"
    Bloquea
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    Fg4.Editable = flexEDKbdMouse
    
    TxtDesc.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar el plan de abastecimiento seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM ges_planabapropro WHERE idpv =" & RstPlanAbas("id") & ""
        xCon.Execute "DELETE * FROM ges_planabadet WHERE idpv =" & RstPlanAbas("id") & ""
        xCon.Execute "DELETE * FROM ges_planaba WHERE id =" & RstPlanAbas("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPlanAbas("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "El plan de abastecimiento se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPlanAbas.Requery
        Dg1.Refresh

    End If
End Sub

Sub Cancelar()
    QueHace = 3
    ActivaTool
    TabOne1.TabEnabled(0) = True
    Bloquea
    Label1.Caption = "Detalle Plan de Abastecimiento"
    TabOne1.CurrTab = 0
End Sub

Sub Bloquea()
    If QueHace <> 3 Then TxtDesc.Locked = False Else TxtDesc.Locked = True
    If QueHace <> 3 Then TxtFchIni.Locked = False Else TxtFchIni.Locked = True
    If QueHace <> 3 Then TxtFchFin.Locked = False Else TxtFchFin.Locked = True
    If QueHace <> 3 Then CmdAdd.Visible = True Else CmdAdd.Visible = False
'    If QueHace <> 3 Then CmdProcesar.Visible = True Else CmdProcesar.Visible = False
'    If QueHace <> 3 Then CmdVerConsolidado.Visible = False Else CmdVerConsolidado.Visible = True
End Sub

Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
End Sub

Sub SetearForm()
    ' POSICIONAMOA EL FORMULARIO
    Me.Caption = "Gestion - Plan de Abastecimiento"
    Me.Width = 12000
    Me.Height = 8200
    
    ' posicionamos el tab
    TabOne1.Left = 0
    TabOne1.Top = 360
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 900
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Eo1.BackColor = &H8000000F
    Eo10.BackColor = &H8000000F
    Eo11.BackColor = &H8000000F
    Eo12.BackColor = &H8000000F
    Eo13.BackColor = &H8000000F
    Eo14.BackColor = &H8000000F
    
    Eo1.BorderWidth = 1
    Eo10.BorderWidth = 1
    Eo11.BorderWidth = 1
    Eo12.BorderWidth = 1
    Eo13.BorderWidth = 1
    Eo14.BorderWidth = 1
    Eo15.BorderWidth = 1
        
    Eo1.ChildSpacing = 1
    Eo10.ChildSpacing = 1
    Eo11.ChildSpacing = 1
    Eo12.ChildSpacing = 1
    Eo13.ChildSpacing = 1
    Eo14.ChildSpacing = 1
    Eo15.ChildSpacing = 1
    
        
    Fg1.BackColor = &HDBFDFD
    Fg2.BackColor = &HDBFDFD
    Fg3.BackColor = &HDBFDFD
    Fg4.BackColor = &HDBFDFD
    
    Fg1.ColWidth(1) = 0
    Fg1.ColWidth(2) = 0
    
    Fg2.ColWidth(1) = 0
    Fg2.ColWidth(2) = 0
    
    Fg3.ColWidth(1) = 0
    Fg3.ColWidth(2) = 0
    
    Fg4.ColWidth(1) = 0
    Fg4.ColWidth(2) = 0
    
        
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H40&
    Fg1.ForeColorSel = &HFFFF&
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShowAndMove
    
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.BackColorSel = &H40&
    Fg2.ForeColorSel = &HFFFF&
    Fg2.AutoSearch = flexSearchFromTop
    Fg2.ExplorerBar = flexExSortShowAndMove
     
    Fg3.SelectionMode = flexSelectionByRow
    Fg3.BackColorSel = &H40&
    Fg3.ForeColorSel = &HFFFF&
    Fg3.AutoSearch = flexSearchFromTop
    Fg3.ExplorerBar = flexExSortShowAndMove
    
    Fg4.SelectionMode = flexSelectionByRow
    Fg4.BackColorSel = &H40&
    Fg4.ForeColorSel = &HFFFF&
    Fg4.AutoSearch = flexSearchFromTop
    Fg4.ExplorerBar = flexExSortShowAndMove
      
    Label1.Width = Eo1.Width - 90
End Sub

Private Sub Form_Load()
    SetearForm
    
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
End Sub

Private Sub Form_Resize()
    CambiarTamaño
End Sub

Sub CambiarTamaño()
    If Me.WindowState = 1 Then Exit Sub
    
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 900
    
    Dg1.Width = Eo1.Width - 60
    Dg1.Height = Eo1.Height - 500
    
    Label1.Width = Eo1.Width - 200
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPlanAbas.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 14 Then
        Set RstPlanAbas = Nothing
        Unload Me
    End If
End Sub

Function Grabar() As Boolean
    If NulosC(TxtDesc.Text) = "" Then
        MsgBox "No ha especificado la descripcion del producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDesc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet2 As New ADODB.Recordset
    Dim RstFue As New ADODB.Recordset
    Dim xId As Double
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM ges_planaba", xCon
        
        xId = HallaCodigoTabla("ges_planaba", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        
        xId = RstPlanAbas("id")
        
        RST_Busq RstCab, "SELECT * FROM ges_planaba WHERE id=" & xId & " ", xCon
        xCon.Execute "DELETE * FROM ges_planabadet WHERE idpv = " & xId & ""
        xCon.Execute "DELETE * FROM ges_planabapropro WHERE idpv = " & xId & ""
        
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM ges_planabadet", xCon
    RST_Busq RstDet2, "SELECT TOP 1 * FROM ges_planabapropro", xCon
    
    RstCab("descripcion") = NulosC(TxtDesc.Text)
    RstCab("fchini") = NulosC(TxtFchIni.Valor)
    RstCab("fchfin") = NulosC(TxtFchFin.Valor)
    RstCab("mesini") = NulosN(Mid(Fg1.TextMatrix(0, 5), 1, 2))
    RstCab("año") = Year(CDate(TxtFchIni.Valor))
    RstCab.Update
    
    Dim xFila, xCol, xMes As Integer
    
    'guardamos los insumos calculados
    'insumos para productos finales
    For xFila = 1 To Fg2.Rows - 1
        For xCol = 5 To Fg2.Cols - 4
            xMes = NulosN(Mid(Fg2.TextMatrix(0, xCol), 1, 2))
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Trim(Fg2.TextMatrix(xFila, 1))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg2.TextMatrix(xFila, xCol))
            RstDet("tipo") = 1
            RstDet.Update
        Next xCol
    Next xFila
    
    'insumos para productos intermedios
    For xFila = 1 To Fg4.Rows - 1
        For xCol = 5 To Fg4.Cols - 4
            xMes = NulosN(Mid(Fg4.TextMatrix(0, xCol), 1, 2))
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Trim(Fg4.TextMatrix(xFila, 1))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg4.TextMatrix(xFila, xCol))
            RstDet("tipo") = 2
            RstDet.Update
        Next xCol
    Next xFila
    
    'grabamos los productos del plan de produccion
    'productos finales
    For xFila = 1 To Fg1.Rows - 1
        For xCol = 5 To Fg1.Cols - 4
            xMes = NulosN(Mid(Fg1.TextMatrix(0, xCol), 1, 2))
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg1.TextMatrix(xFila, 1))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = NulosN(Fg1.TextMatrix(xFila, xCol))
            RstDet2("tipo") = 1
            RstDet2.Update
        Next xCol
    Next xFila
    
    'productos intermedios
    For xFila = 1 To Fg3.Rows - 1
        For xCol = 5 To Fg3.Cols - 4
            xMes = NulosN(Mid(Fg3.TextMatrix(0, xCol), 1, 2))
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg3.TextMatrix(xFila, 1))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = Fg3.TextMatrix(xFila, xCol)
            RstDet2("tipo") = 2
            RstDet2.Update
        Next xCol
    Next xFila
       
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
       
    xCon.CommitTrans
    MsgBox "El plan de abastecimiento se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Nuevo()
    QueHace = 1
    xHorIni = Time

    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label1.Caption = "Agregando Plan de Abastecimiento"
    Bloquea
    Blanquea
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg3.Rows = 1
    Fg4.Rows = 1
    
    TxtDesc.SetFocus
End Sub

Sub MuestraSegundoTab()
    Dim xSQL As String
    Dim xMes As String
    Dim xRst As New ADODB.Recordset
    Dim A, B, xMesIni, xAñoTra As Integer
    Dim xStock, xDiferencia As Double

    Bloquea
    TxtDesc.Text = RstPlanAbas("descripcion")
    TxtFchIni.Valor = Format(RstPlanAbas("fchini"), "dd/mm/yyyy")
    TxtFchFin.Valor = Format(RstPlanAbas("fchfin"), "dd/mm/yyyy")
    xMesIni = NulosN(RstPlanAbas("mesini"))
    xAñoTra = NulosN(RstPlanAbas("año"))
    
    ' MOSTRAMOS LOS PRODUCTOS TERMINADOS
    xSQL = "TRANSFORM Sum(ges_planabapropro.cantidad) AS SumaDecantidad " _
        & " SELECT ges_planabapropro.idpv, ges_planabapropro.tipo, ges_planabapropro.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planabapropro.cantidad) AS [total] " _
        & " FROM ges_planabapropro LEFT JOIN (alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON ges_planabapropro.codpro = alm_inventario.id " _
        & " Where (((ges_planabapropro.idpv) = " & RstPlanAbas("id") & ") And ((ges_planabapropro.tipo) = 1)) " _
        & " GROUP BY ges_planabapropro.idpv, ges_planabapropro.tipo, ges_planabapropro.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabapropro.idmes"

    RST_Busq xRst, xSQL, xCon
    
    Fg1.Cols = 5
    Fg1.Rows = 1
    Fg3.Cols = 5
    Fg3.Rows = 1
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg1.Cols = Fg1.Cols + 1
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg1.Cols = Fg1.Cols + 1
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Programado"
        Fg1.ColWidth(Fg1.Cols - 1) = 1100
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Stock"
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Diferencia"
        Fg1.ColWidth(Fg1.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = xRst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRst("abrev")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg1.TextMatrix(0, B), 1, 2))))
                Fg1.TextMatrix(Fg1.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
            
            xStock = SaldoActual(NulosN(Fg1.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 3) = Format(xRst("total"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 2) = Format(xStock, "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 1) = Format(xRst("total") - xStock, "0.00")
    
            xRst.MoveNext
        Next A
    End If
   
    ' ***********************************
    ' MOSTRAMOS LOS PRODUCTOS INTERMEDIOS
    xSQL = "TRANSFORM Sum(ges_planabapropro.cantidad) AS SumaDecantidad " _
        & " SELECT ges_planabapropro.idpv, ges_planabapropro.tipo, ges_planabapropro.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planabapropro.cantidad) AS [total] " _
        & " FROM ges_planabapropro LEFT JOIN (alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON ges_planabapropro.codpro = alm_inventario.id " _
        & " Where (((ges_planabapropro.idpv) = " & RstPlanAbas("id") & ") And ((ges_planabapropro.tipo) = 2)) " _
        & " GROUP BY ges_planabapropro.idpv, ges_planabapropro.tipo, ges_planabapropro.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabapropro.idmes"

    RST_Busq xRst, xSQL, xCon
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg3.Cols = Fg3.Cols + 1
            Fg3.TextMatrix(0, Fg3.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg3.Cols = Fg3.Cols + 1
            Fg3.TextMatrix(0, Fg3.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Programado"
        Fg3.ColWidth(Fg3.Cols - 1) = 1100
        
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Stock"
        
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Diferencia"
        Fg3.ColWidth(Fg3.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = xRst("codpro")
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = xRst("idunimed")
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = xRst("descripcion")
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = xRst("abrev")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg3.TextMatrix(0, B), 1, 2))))
                Fg3.TextMatrix(Fg3.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
            
            xStock = SaldoActual(NulosN(Fg3.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 3) = Format(xRst("total"), "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 2) = Format(xStock, "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 1) = Format(xRst("total") - xStock, "0.00")
    
            xRst.MoveNext
        Next A
    End If
   
   
   
   
   
   
    ' **************************************
    ' MOSTRAMOS LOS INSUMOS DE LOS TEMINADOS
    xSQL = "TRANSFORM Sum(ges_planabadet.cantidad) AS SumaDecantidad " _
        & " SELECT ges_planabadet.idpv, ges_planabadet.tipo, ges_planabadet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planabadet.cantidad) AS [total] " _
        & " FROM (ges_planabadet LEFT JOIN alm_inventario ON ges_planabadet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & " Where (((ges_planabadet.idpv) = " & RstPlanAbas("id") & ") And ((ges_planabadet.tipo) = 1)) " _
        & " GROUP BY ges_planabadet.idpv, ges_planabadet.tipo, ges_planabadet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabadet.idmes"

    Fg2.Cols = 5
    Fg2.Rows = 1
    Fg4.Cols = 5
    Fg4.Rows = 1
    
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg2.Cols = Fg2.Cols + 1
            Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg2.Cols = Fg2.Cols + 1
            Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Programado"
        Fg2.ColWidth(Fg2.Cols - 1) = 1100
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Stock"
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Diferencia"
        Fg2.ColWidth(Fg2.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRst("codpro")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRst("idunimed")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = xRst("descripcion")
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = xRst("abrev")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))))
                Fg2.TextMatrix(Fg2.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
                                   
            xStock = SaldoActual(NulosN(Fg2.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 3) = Format(xRst("total"), "0.00")
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 2) = Format(xStock, "0.00")
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 1) = Format(xRst("total") - xStock, "0.00")
            
            xRst.MoveNext
        Next A
    End If
    
    
    ' ****************************************
    ' MOSTRAMOS LOS INSUMOS DE LOS INTERMEDIOS
    xSQL = "TRANSFORM Sum(ges_planabadet.cantidad) AS SumaDecantidad " _
        & " SELECT ges_planabadet.idpv, ges_planabadet.tipo, ges_planabadet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planabadet.cantidad) AS [total] " _
        & " FROM (ges_planabadet LEFT JOIN alm_inventario ON ges_planabadet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & " Where (((ges_planabadet.idpv) = " & RstPlanAbas("id") & ") And ((ges_planabadet.tipo) = 2)) " _
        & " GROUP BY ges_planabadet.idpv, ges_planabadet.tipo, ges_planabadet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabadet.idmes"

    
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        For A = xMesIni To 12
            Fg4.Cols = Fg4.Cols + 1
            Fg4.TextMatrix(0, Fg4.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg4.Cols = Fg4.Cols + 1
            Fg4.TextMatrix(0, Fg4.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Programado"
        Fg4.ColWidth(Fg4.Cols - 1) = 1100
        
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Stock"
        
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Diferencia"
        Fg4.ColWidth(Fg4.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(Fg4.Rows - 1, 1) = xRst("codpro")
            Fg4.TextMatrix(Fg4.Rows - 1, 2) = xRst("idunimed")
            Fg4.TextMatrix(Fg4.Rows - 1, 3) = xRst("descripcion")
            Fg4.TextMatrix(Fg4.Rows - 1, 4) = xRst("abrev")
            
            ' ESCRIBIMOS LOS MESES
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg4.TextMatrix(0, B), 1, 2))))
                Fg4.TextMatrix(Fg4.Rows - 1, B) = Format(xRst(xMes), "0.00")
            Next B
                                   
            xStock = SaldoActual(NulosN(Fg4.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 3) = Format(xRst("total"), "0.00")
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 2) = Format(xStock, "0.00")
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 1) = Format(xRst("total") - xStock, "0.00")
            
            xRst.MoveNext
        Next A
    End If
    TabOne2.CurrTab = 0
    PintarGrid
End Sub

