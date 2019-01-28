VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanAbastecimiento3 
   Caption         =   "Compras - Plan de Abastecimiento"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBarra 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   885
      Left            =   3360
      TabIndex        =   42
      Top             =   4380
      Visible         =   0   'False
      Width           =   6180
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   90
         TabIndex        =   43
         Top             =   390
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label LblBarra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Productos Terminados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   90
         Width           =   2970
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   6165
         X2              =   6165
         Y1              =   15
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6150
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6165
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   270
         Left            =   60
         Top             =   60
         Width           =   6075
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   8730
      Left            =   165
      TabIndex        =   0
      Top             =   360
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
         TabIndex        =   1
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
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   90
            Width           =   12195
         End
      End
      Begin SizerOneLibCtl.ElasticOne Eo10 
         Height          =   8310
         Left            =   45
         TabIndex        =   4
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
            Height          =   6900
            Left            =   90
            TabIndex        =   14
            Top             =   1320
            Width           =   12195
            _cx             =   21511
            _cy             =   12171
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
            Caption         =   "   &Terminados   |   &Intermedios  |   &Total  "
            Align           =   0
            CurrTab         =   2
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
               Height          =   6480
               Left            =   -12750
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   45
               Width           =   12105
               _cx             =   21352
               _cy             =   11430
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
                  Height          =   3015
                  Left            =   90
                  TabIndex        =   22
                  Top             =   3375
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5318
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
                  Cols            =   6
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
                  Height          =   2955
                  Left            =   90
                  TabIndex        =   21
                  Top             =   90
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5212
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":0189
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
                  TabIndex        =   18
                  TabStop         =   0   'False
                  Top             =   3105
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
                  _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0232
                  Begin VB.CommandButton Command4 
                     Height          =   30
                     Left            =   5955
                     Picture         =   "FrmPlanAbastecimiento3.frx":0272
                     Style           =   1  'Graphical
                     TabIndex        =   25
                     Top             =   90
                     Width           =   5880
                  End
                  Begin VB.CommandButton Command2 
                     Height          =   30
                     Left            =   90
                     Picture         =   "FrmPlanAbastecimiento3.frx":03B0
                     Style           =   1  'Graphical
                     TabIndex        =   24
                     Top             =   90
                     Width           =   5805
                  End
               End
            End
            Begin SizerOneLibCtl.ElasticOne Eo14 
               Height          =   6480
               Left            =   -13050
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   45
               Width           =   12105
               _cx             =   21352
               _cy             =   11430
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
               _GridInfo       =   $"FrmPlanAbastecimiento3.frx":04EE
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   2985
                  Left            =   90
                  TabIndex        =   20
                  Top             =   3405
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5265
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":053F
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
                  TabIndex        =   19
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":05E6
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
                  Height          =   6300
                  Left            =   90
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   11113
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
                  _GridInfo       =   $"FrmPlanAbastecimiento3.frx":068F
                  Begin VB.CommandButton Command3 
                     Height          =   6120
                     Left            =   5955
                     Picture         =   "FrmPlanAbastecimiento3.frx":06D4
                     Style           =   1  'Graphical
                     TabIndex        =   26
                     Top             =   90
                     Width           =   5880
                  End
                  Begin VB.CommandButton Command1 
                     Height          =   6120
                     Left            =   90
                     Picture         =   "FrmPlanAbastecimiento3.frx":0812
                     Style           =   1  'Graphical
                     TabIndex        =   23
                     Top             =   90
                     Width           =   5805
                  End
               End
            End
            Begin SizerOneLibCtl.ElasticOne Eo16 
               Height          =   6480
               Left            =   45
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   45
               Width           =   12105
               _cx             =   21352
               _cy             =   11430
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
               BackColor       =   12632319
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
               _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0950
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   3015
                  Left            =   90
                  TabIndex        =   28
                  Top             =   3375
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5318
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":09A1
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
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   2955
                  Left            =   90
                  TabIndex        =   29
                  Top             =   90
                  Width           =   11925
                  _cx             =   21034
                  _cy             =   5212
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanAbastecimiento3.frx":0A48
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
               Begin SizerOneLibCtl.ElasticOne Eo17 
                  Height          =   210
                  Left            =   90
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   3105
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
                  _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0AF1
                  Begin VB.CommandButton Command6 
                     Height          =   30
                     Left            =   90
                     Picture         =   "FrmPlanAbastecimiento3.frx":0B31
                     Style           =   1  'Graphical
                     TabIndex        =   32
                     Top             =   90
                     Width           =   5805
                  End
                  Begin VB.CommandButton Command5 
                     Height          =   30
                     Left            =   5955
                     Picture         =   "FrmPlanAbastecimiento3.frx":0C6F
                     Style           =   1  'Graphical
                     TabIndex        =   31
                     Top             =   90
                     Width           =   5880
                  End
               End
            End
         End
         Begin SizerOneLibCtl.ElasticOne Eo13 
            Height          =   870
            Left            =   90
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   390
            Width           =   12195
            _cx             =   21511
            _cy             =   1535
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
            _GridInfo       =   $"FrmPlanAbastecimiento3.frx":0DAD
            Begin VB.Frame Frame3 
               BackColor       =   &H0080C0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   690
               Left            =   10035
               TabIndex        =   36
               Top             =   90
               Width           =   2070
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
                  Height          =   195
                  Left            =   1065
                  TabIndex        =   40
                  Top             =   510
                  Width           =   990
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Registros : "
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   30
                  TabIndex        =   39
                  Top             =   510
                  Width           =   1020
               End
               Begin VB.Shape Shape4 
                  BackColor       =   &H000000C0&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H00C0C0C0&
                  Height          =   180
                  Left            =   0
                  Top             =   270
                  Width           =   540
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "= Item sin Stock"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   600
                  TabIndex        =   38
                  Top             =   270
                  Width           =   1395
               End
               Begin VB.Shape Shape3 
                  BackColor       =   &H00C00000&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H00C0C0C0&
                  Height          =   180
                  Left            =   0
                  Top             =   60
                  Width           =   540
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "= Item con Stock"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   600
                  TabIndex        =   37
                  Top             =   120
                  Width           =   1470
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   690
               Left            =   6945
               TabIndex        =   33
               Top             =   90
               Width           =   3030
               Begin VB.CommandButton CmdAdd 
                  Caption         =   "Agregar Plan de Producción"
                  Height          =   525
                  Left            =   1275
                  TabIndex        =   35
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.CommandButton CmdVerEst 
                  Caption         =   "&Ver Estacionalidad"
                  Height          =   525
                  Left            =   75
                  TabIndex        =   34
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   690
               Left            =   90
               TabIndex        =   6
               Top             =   90
               Width           =   6795
               Begin VB.TextBox TxtDesc 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   945
                  Locked          =   -1  'True
                  TabIndex        =   7
                  Text            =   "TxtDesc"
                  Top             =   75
                  Width           =   5775
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
                  Height          =   300
                  Left            =   945
                  TabIndex        =   8
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
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Término"
                  Height          =   195
                  Left            =   3465
                  TabIndex        =   12
                  Top             =   420
                  Width           =   930
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripción"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   11
                  Top             =   105
                  Width           =   840
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Inicio"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   10
                  Top             =   420
                  Width           =   735
               End
            End
         End
         Begin VB.Label Label2 
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
            Height          =   240
            Left            =   90
            TabIndex        =   13
            Top             =   90
            Width           =   12195
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10
      Top             =   90
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
            Picture         =   "FrmPlanAbastecimiento3.frx":0DFD
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":1341
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":16D3
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":1857
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":1CAB
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":1DC3
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":2307
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":284B
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":295F
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":2A73
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":2EC7
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":3033
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":357B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento3.frx":3895
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   12825
      _ExtentX        =   22622
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar plan de Abastecimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar plan de Abastecimiento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar plan de Abastecimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar plan de Abastecimiento"
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
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Plan de produccion productos terminados"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Plan de produccion de produccion productois"
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
Dim xTipoGrilla As Integer '--Indica la grilla activada, esto indicara que la grilla se prodra activar
                           '--1: Productos Teminados - del plan de produccion
                           '--2: Insumos/Materia Prima para Productos Terminados
                           '--3: Productos Teminados - del plan de produccion
                           '--4: Insumos/Materia Prima para Productos Intermedios
                           '--5: Resumen de los Productos terminados con Intermedios
                           '--6: Resumen de los Insumos o Materia Prima de los Productos Terminados con Intermedios

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
    
    xform.Titulo = "Buscando Plan de Producción"
    
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
        
        '----------------------------------------------------
        FraBarra.Visible = True
        FraBarra.Left = 3360
        FraBarra.Top = 4380
        
        ProgressBar1.Value = 1
        ProgressBar1.Min = 1
        '----------------------------------------------------
        
        
        MostrarTerminados xId, xRs("mesini"), xRs("año")
        MostrarIntermedios xId, xRs("mesini"), xRs("año")
        
        '--mostrar los acumulados
        LblBarra.Caption = "Procesando Resumen de Productos"
        MostrarAcumulado Fg5, Fg1, "T", 1, False, TxtFchIni.Valor, ProgressBar1
        MostrarAcumulado Fg5, Fg3, "I", 1, True, TxtFchIni.Valor, ProgressBar1
        
        LblBarra.Caption = "Procesando Resumen de Materia Prima/ Insumos"
        MostrarAcumulado Fg6, Fg2, "T", 2, False, TxtFchIni.Valor, ProgressBar1
        MostrarAcumulado Fg6, Fg4, "I", 2, True, TxtFchIni.Valor, ProgressBar1
        '--------
        FraBarra.Visible = False
        
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
        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet.idmes in (1,2,3,4,5,6,7,8,9,10,11,12) "
    
    RST_Busq xRst, xSQL, xCon
       
    If xRst.RecordCount <> 0 Then
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Productos Terminados"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
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
        
''        Fg1.Cols = Fg1.Cols + 1
''        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Stock"
''
''        Fg1.Cols = Fg1.Cols + 1
''        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Diferencia"
''        Fg1.ColWidth(Fg1.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = ""
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRst("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(xRst("abrev"))
            
            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg1.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg1.TextMatrix(Fg1.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B
            
''            xStock = SaldoActual(NulosN(Fg1.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 2) = Format(xStock, "0.00")
''            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 1) = Format(xRst("total") - xStock, "0.00")
    
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
        & " GROUP BY terminados.ins_iditem, terminados.ins_desc, terminados.ins_idunimed, terminados.ins_unimed PIVOT terminados.idmes in (1,2,3,4,5,6,7,8,9,10,11,12)"

    
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Insumos de Productos Terminados"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
    
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
        
''        Fg2.Cols = Fg2.Cols + 1
''        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Stock"
''
''        Fg2.Cols = Fg2.Cols + 1
''        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Diferencia"
''        Fg2.ColWidth(Fg2.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRst("ins_iditem")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRst("ins_idunimed")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = ""
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(xRst("ins_desc"))
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosC(xRst("ins_unimed"))
            
            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg2.TextMatrix(Fg2.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B
                                   
            xStock = SaldoActual(NulosN(Fg2.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
            
''            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 2) = Format(xStock, "0.00")
''            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 1) = Format(xRst("total") - xStock, "0.00")
            
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
        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet2.idmes in (1,2,3,4,5,6,7,8,9,10,11,12)"
    
    RST_Busq xRst, xSQL, xCon
    
    If xRst.RecordCount <> 0 Then
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Productos Intermedios"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
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
        
''        Fg3.Cols = Fg3.Cols + 1
''        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Stock"
''
''        Fg3.Cols = Fg3.Cols + 1
''        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Diferencia"
''        Fg3.ColWidth(Fg3.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = xRst("codpro")
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = xRst("idunimed")
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = ""
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosC(xRst("descripcion"))
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosC(xRst("abrev"))
            
            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg3.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg3.TextMatrix(Fg3.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B
            
''            xStock = SaldoActual(NulosN(Fg3.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 2) = Format(xStock, "0.00")
''            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 1) = Format(xRst("total") - xStock, "0.00")
''
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
        & " GROUP BY intermedio.ins_iditem, intermedio.ins_desc, intermedio.ins_idunimed, intermedio.ins_unimed ORDER BY intermedio.ins_desc PIVOT intermedio.idmes in (1,2,3,4,5,6,7,8,9,10,11,12)"
   
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Insumos de Productos Intermedios"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
        xRst.MoveFirst
        '--configurar los periodos
        For A = xMesIni To 12
            Fg4.Cols = Fg4.Cols + 1
            Fg4.TextMatrix(0, Fg4.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
        Next A
        
        For A = 1 To xMesIni - 1
            Fg4.Cols = Fg4.Cols + 1
            Fg4.TextMatrix(0, Fg4.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra + 1, "0000")
        Next A
        '--configurar el total programado
        Fg4.Cols = Fg4.Cols + 1
        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Programado"
        Fg4.ColWidth(Fg4.Cols - 1) = 1100
        
''        Fg4.Cols = Fg4.Cols + 1
''        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Stock"
''
''        Fg4.Cols = Fg4.Cols + 1
''        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Diferencia"
''        Fg4.ColWidth(Fg4.Cols - 1) = 1100
        
        For A = 1 To xRst.RecordCount
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
            Fg4.Rows = Fg4.Rows + 1
            '--datos principales del item
            Fg4.TextMatrix(Fg4.Rows - 1, 1) = xRst("ins_iditem")
            Fg4.TextMatrix(Fg4.Rows - 1, 2) = xRst("ins_idunimed")
            Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
            Fg4.TextMatrix(Fg4.Rows - 1, 4) = NulosC(xRst("ins_desc"))
            Fg4.TextMatrix(Fg4.Rows - 1, 5) = NulosC(xRst("ins_unimed"))
            
            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg4.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg4.TextMatrix(Fg4.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B
                                   
''            xStock = SaldoActual(NulosN(Fg4.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)

            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 2) = Format(xStock, "0.00")
''            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 1) = Format(xRst("total") - xStock, "0.00")
            
            xRst.MoveNext
        Next A
    End If
    
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Fg1_Click()
    xTipoGrilla = 1
    LblNumReg.Caption = Format(Fg1.Rows - 1, "000")
End Sub

Private Sub Fg2_Click()
    xTipoGrilla = 2
    LblNumReg.Caption = Format(Fg2.Rows - 1, "000")
End Sub

Private Sub Fg3_Click()
    xTipoGrilla = 3
    LblNumReg.Caption = Format(Fg3.Rows - 1, "000")
End Sub

Private Sub Fg4_Click()
    xTipoGrilla = 4
    LblNumReg.Caption = Format(Fg4.Rows - 1, "000")
End Sub

Private Sub Fg5_Click()
    xTipoGrilla = 5
    LblNumReg.Caption = Format(Fg5.Rows - 1, "000")
End Sub

Private Sub Fg6_Click()
    xTipoGrilla = 6
    LblNumReg.Caption = Format(Fg6.Rows - 1, "000")
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
'    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 3, Fg1.Rows - 1, Fg1.Cols - 3, &HFFFFC0, flexFillRepeat
'    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 2, Fg1.Rows - 1, Fg1.Cols - 2, &HC0FFC0, flexFillRepeat

'    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 3, Fg2.Rows - 1, Fg2.Cols - 3, &HFFFFC0, flexFillRepeat
'    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 2, Fg2.Rows - 1, Fg2.Cols - 2, &HC0FFC0, flexFillRepeat
'
'    GRID_COLOR_FONDO Fg3, 1, Fg3.Cols - 3, Fg3.Rows - 1, Fg3.Cols - 3, &HFFFFC0, flexFillRepeat
'    GRID_COLOR_FONDO Fg3, 1, Fg3.Cols - 2, Fg3.Rows - 1, Fg3.Cols - 2, &HC0FFC0, flexFillRepeat
'
'    GRID_COLOR_FONDO Fg4, 1, Fg4.Cols - 3, Fg4.Rows - 1, Fg4.Cols - 3, &HFFFFC0, flexFillRepeat
'    GRID_COLOR_FONDO Fg4, 1, Fg4.Cols - 2, Fg4.Rows - 1, Fg4.Cols - 2, &HC0FFC0, flexFillRepeat

    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 1, Fg2.Rows - 1, Fg2.Cols - 1, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg3, 1, Fg3.Cols - 1, Fg3.Rows - 1, Fg3.Cols - 1, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg4, 1, Fg4.Cols - 1, Fg4.Rows - 1, Fg4.Cols - 1, &HFFFFC0, flexFillRepeat

        
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
    
    '--total productos
    GRID_COLOR_FONDO Fg5, 1, Fg5.Cols - 5, Fg5.Rows - 1, Fg5.Cols - 5, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg5, 1, Fg5.Cols - 4, Fg5.Rows - 1, Fg5.Cols - 2, &HC0FFC0, flexFillRepeat
    '--total insumos
    GRID_COLOR_FONDO Fg6, 1, Fg6.Cols - 5, Fg6.Rows - 1, Fg6.Cols - 5, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg6, 1, Fg6.Cols - 4, Fg6.Rows - 1, Fg6.Cols - 2, &HC0FFC0, flexFillRepeat

    For A = 1 To Fg5.Cols - 1
        Fg5.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg5, 0, A, , True, &H8000000F, Fg5.TextMatrix(0, A)
    Next A
    
    For A = 1 To Fg4.Cols - 1
        Fg6.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg6, 0, A, , True, &H8000000F, Fg6.TextMatrix(0, A)
    Next A
    
    Fg1.FrozenCols = 5
    Fg2.FrozenCols = 5
    Fg3.FrozenCols = 5
    Fg4.FrozenCols = 5
    Fg5.FrozenCols = 6
    Fg6.FrozenCols = 6
    
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
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
    LblNumReg.Caption = 0
    
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    Fg1.Cols = 6
    Fg1.Rows = 1
    
    Fg3.Cols = 6
    Fg3.Rows = 1
    
    Fg2.Cols = 6
    Fg2.Rows = 1
    
    Fg4.Cols = 6
    Fg4.Rows = 1
    
    Fg5.Cols = 6
    Fg5.Rows = 1
    Fg6.Cols = 6
    Fg6.Rows = 1
    
    DoEvents
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
    Frame3.BackColor = &H8000000F
    
    Eo1.BackColor = &H8000000F
    Eo10.BackColor = &H8000000F
    Eo11.BackColor = &H8000000F
    Eo12.BackColor = &H8000000F
    Eo13.BackColor = &H8000000F
    Eo14.BackColor = &H8000000F
    Eo16.BackColor = &H8000000F
    Eo17.BackColor = &H8000000F
    
    Eo1.BorderWidth = 1
    Eo10.BorderWidth = 1
    Eo11.BorderWidth = 1
    Eo12.BorderWidth = 1
    Eo13.BorderWidth = 1
    Eo14.BorderWidth = 1
    Eo15.BorderWidth = 1
    Eo16.BorderWidth = 1
    Eo17.BorderWidth = 1
        
    Eo1.ChildSpacing = 1
    Eo10.ChildSpacing = 1
    Eo11.ChildSpacing = 1
    Eo12.ChildSpacing = 1
    Eo13.ChildSpacing = 1
    Eo14.ChildSpacing = 1
    Eo15.ChildSpacing = 1
    Eo16.ChildSpacing = 1
    Eo17.ChildSpacing = 1
        
    Fg1.BackColor = &HDBFDFD
    Fg2.BackColor = &HDBFDFD
    Fg3.BackColor = &HDBFDFD
    Fg4.BackColor = &HDBFDFD
    Fg5.BackColor = &HDBFDFD
    Fg6.BackColor = &HDBFDFD
    
    Fg1.ColWidth(1) = 0
    Fg1.ColWidth(2) = 0
    Fg1.ColWidth(3) = 0
    
    Fg2.ColWidth(1) = 0
    Fg2.ColWidth(2) = 0
    Fg2.ColWidth(3) = 0
    
    Fg3.ColWidth(1) = 0
    Fg3.ColWidth(2) = 0
    Fg3.ColWidth(3) = 0
    
    Fg4.ColWidth(1) = 0
    Fg4.ColWidth(2) = 0
    Fg4.ColWidth(3) = 0
    
    Fg5.ColWidth(1) = 0
    Fg5.ColWidth(2) = 0
    Fg5.ColWidth(3) = 0
    
    Fg6.ColWidth(1) = 0
    Fg6.ColWidth(2) = 0
    Fg6.ColWidth(3) = 0
    
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
      
    Fg5.SelectionMode = flexSelectionByRow
    Fg5.BackColorSel = &H40&
    Fg5.ForeColorSel = &HFFFF&
    Fg5.AutoSearch = flexSearchFromTop
    Fg5.ExplorerBar = flexExSortShowAndMove
    
    Fg6.SelectionMode = flexSelectionByRow
    Fg6.BackColorSel = &H40&
    Fg6.ForeColorSel = &HFFFF&
    Fg6.AutoSearch = flexSearchFromTop
    Fg6.ExplorerBar = flexExSortShowAndMove
      
    Label1.Width = Eo1.Width - 90
End Sub

Private Sub Form_Load()
    SetearForm
    xTipoGrilla = 0
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
    
    If Button.Index = 8 Then ExportarExcel
        
    If Button.Index = 15 Then
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
    RstCab("mesini") = Month(CDate(TxtFchIni.Valor))
    '--RstCab("mesini") = NulosN(Mid(Fg1.TextMatrix(0, 5), 1, 2))
    RstCab("año") = Year(CDate(TxtFchIni.Valor))
    RstCab.Update
    
    Dim xFila, xCol, xMes As Integer
    
    'guardamos los insumos calculados
    'insumos para productos finales
    For xFila = 1 To Fg2.Rows - 1
        For xCol = 6 To Fg2.Cols - 2
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
        For xCol = 6 To Fg4.Cols - 2
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
        For xCol = 6 To Fg1.Cols - 2
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
        For xCol = 6 To Fg3.Cols - 2
            xMes = NulosN(Mid(Fg3.TextMatrix(0, xCol), 1, 2))
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg3.TextMatrix(xFila, 1))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = NulosN(Fg3.TextMatrix(xFila, xCol))
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
    
    Fg5.Rows = 1
    Fg6.Rows = 1
    
    
    TxtDesc.SetFocus
End Sub

Sub MuestraSegundoTab()
    Dim xSQL As String
    Dim xMes As String
    Dim xRst As New ADODB.Recordset
    Dim A, B, xMesIni, xAñoTra As Integer
    Dim xStock, xDiferencia As Double
    
    Blanquea
    
    Bloquea
    
    '----------------------------------------------------
    FraBarra.Visible = True
    FraBarra.Left = 3360
    FraBarra.Top = 4380
    
    ProgressBar1.Value = 1
    ProgressBar1.Min = 1
    '----------------------------------------------------
    
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
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabapropro.idmes in (1,2,3,4,5,6,7,8,9,10,11,12)"

    RST_Busq xRst, xSQL, xCon
    
    If xRst.RecordCount <> 0 Then
    
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Productos Terminados"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
        
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
        
''        Fg1.Cols = Fg1.Cols + 1
''        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Stock"
''
''        Fg1.Cols = Fg1.Cols + 1
''        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Diferencia"
''        Fg1.ColWidth(Fg1.Cols - 1) = 1100
''
        For A = 1 To xRst.RecordCount
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = ""
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = xRst("abrev")

            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg1.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg1.TextMatrix(Fg1.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B

''            xStock = SaldoActual(NulosN(Fg1.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
              Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 2) = Format(xStock, "0.00")
''            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.Cols - 1) = Format(xRst("total") - xStock, "0.00")

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
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabapropro.idmes  in (1,2,3,4,5,6,7,8,9,10,11,12) "

    RST_Busq xRst, xSQL, xCon
    
    If xRst.RecordCount <> 0 Then
        
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Productos Intermedios"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
        
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
        
''        Fg3.Cols = Fg3.Cols + 1
''        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Stock"
''
''        Fg3.Cols = Fg3.Cols + 1
''        Fg3.TextMatrix(0, Fg3.Cols - 1) = "Diferencia"
''        Fg3.ColWidth(Fg3.Cols - 1) = 1100
''
        For A = 1 To xRst.RecordCount
            
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
        
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = xRst("codpro")
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = xRst("idunimed")
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = ""
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = xRst("descripcion")
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = xRst("abrev")

            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg3.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg3.TextMatrix(Fg3.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B

''            xStock = SaldoActual(NulosN(Fg3.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 2) = Format(xStock, "0.00")
''            Fg3.TextMatrix(Fg3.Rows - 1, Fg3.Cols - 1) = Format(xRst("total") - xStock, "0.00")

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
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabadet.idmes  in (1,2,3,4,5,6,7,8,9,10,11,12) "

    
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Insumos de Productos Terminados"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
        
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
        
''        Fg2.Cols = Fg2.Cols + 1
''        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Stock"
''
''        Fg2.Cols = Fg2.Cols + 1
''        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Diferencia"
''        Fg2.ColWidth(Fg2.Cols - 1) = 1100
''
        For A = 1 To xRst.RecordCount
        
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
        
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRst("codpro")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRst("idunimed")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = ""
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(xRst("descripcion"))
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosC(xRst("abrev"))

            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg2.TextMatrix(Fg2.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B

''            xStock = SaldoActual(NulosN(Fg2.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 2) = Format(xStock, "0.00")
''            Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Cols - 1) = Format(xRst("total") - xStock, "0.00")

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
        & " ORDER BY alm_inventario.descripcion PIVOT ges_planabadet.idmes  in (1,2,3,4,5,6,7,8,9,10,11,12) "

    
    RST_Busq xRst, xSQL, xCon
   
    If xRst.RecordCount <> 0 Then
        '----------------------------------------------------
        LblBarra.Caption = "Procesando Insumos de Productos Intermedios"
        ProgressBar1.Max = xRst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        '----------------------------------------------------
        
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
        
''        Fg4.Cols = Fg4.Cols + 1
''        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Stock"
''
''        Fg4.Cols = Fg4.Cols + 1
''        Fg4.TextMatrix(0, Fg4.Cols - 1) = "Diferencia"
''        Fg4.ColWidth(Fg4.Cols - 1) = 1100
''
        For A = 1 To xRst.RecordCount
            
            '----------------------------------------------------
            ProgressBar1.Value = A
            DoEvents
            '----------------------------------------------------
            
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(Fg4.Rows - 1, 1) = xRst("codpro")
            Fg4.TextMatrix(Fg4.Rows - 1, 2) = xRst("idunimed")
            Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
            Fg4.TextMatrix(Fg4.Rows - 1, 4) = NulosC(xRst("descripcion"))
            Fg4.TextMatrix(Fg4.Rows - 1, 5) = NulosC(xRst("abrev"))

            ' ESCRIBIMOS LOS MESES
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg4.TextMatrix(0, B), 1, 2))))
                If xMes <> 0 Then Fg4.TextMatrix(Fg4.Rows - 1, B) = Format(xRst(xMes), FORMAT_MONTO)
            Next B

''            xStock = SaldoActual(NulosN(Fg4.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 1) = Format(xRst("total"), FORMAT_MONTO)
''            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 3) = Format(xRst("total"), "0.00")
''            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 2) = Format(xStock, "0.00")
''            Fg4.TextMatrix(Fg4.Rows - 1, Fg4.Cols - 1) = Format(xRst("total") - xStock, "0.00")

            xRst.MoveNext
        Next A
    End If
    TabOne2.CurrTab = 0
    
    '--mostrar los acumulados
    
    '--(xTipo)   T = Terminados; I = Intermedios
    '--(xOrigen) 1 = Productos; 2 = Insumos o Materia Prima
    '--(xDatosAdicionales) False = No se muestran; True = Si se muestran los datos (Stock Ini, Producido o Comprado, Total, Diferencia)
    '--xFechaIni = Indica la fecha incial
    LblBarra.Caption = "Procesando Resumen de Productos"
    MostrarAcumulado Fg5, Fg1, "T", 1, False, TxtFchIni.Valor, ProgressBar1
    MostrarAcumulado Fg5, Fg3, "I", 1, True, TxtFchIni.Valor, ProgressBar1
    
    LblBarra.Caption = "Procesando Resumen de Materia Prima/ Insumos"
    MostrarAcumulado Fg6, Fg2, "T", 2, False, TxtFchIni.Valor, ProgressBar1
    MostrarAcumulado Fg6, Fg4, "I", 2, True, TxtFchIni.Valor, ProgressBar1
    '--------
    
    FraBarra.Visible = False
    
    PintarGrid
End Sub






Private Sub ExportarExcel()
    Dim xTitulo As String
    Dim xPeriodo As String
    Dim xFg As VSFlexGrid
    On Error GoTo LaCague
    
    If IsDate(TxtFchIni.Valor) = False Then
        MsgBox "Falta especificar la fecha de inicio", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If IsDate(TxtFchFin.Valor) = False Then
        MsgBox "Falta especificar la fecha final", vbInformation, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    xPeriodo = "Del: " & TxtFchIni.Valor & " Al: " & TxtFchFin.Valor
    
    If TabOne2.CurrTab = 0 Then
        If xTipoGrilla = 1 Then
            xTitulo = "Plan de Producción de Productos Terminados"
            Set xFg = Fg1
        ElseIf xTipoGrilla = 2 Then
            xTitulo = "Plan de Abastecimiento de Productos Terminados"
            Set xFg = Fg2
        Else
            Exit Sub
        End If
    ElseIf TabOne2.CurrTab = 1 Then
        If xTipoGrilla = 3 Then
            xTitulo = "Plan de Producción de Productos Intermedios"
            Set xFg = Fg3
        ElseIf xTipoGrilla = 4 Then
            xTitulo = "Plan de Abastecimiento de Productos Intermedios"
            Set xFg = Fg4
        Else
            Exit Sub
        End If
    ElseIf TabOne2.CurrTab = 2 Then
        If xTipoGrilla = 5 Then
            xTitulo = "Resumen del Plan de Producción "
            Set xFg = Fg5
        ElseIf xTipoGrilla = 6 Then
            xTitulo = "Resumen del Plan de Abastecimiento "
            Set xFg = Fg6
        Else
            Exit Sub
        End If
    End If
    
    Dim xExport As New SGI2_funciones.Formularios
    xExport.VSFlexGrid_Exportar_MSExcel xCon, xFg, xTitulo, xPeriodo, "", xTitulo
    Set xExport = Nothing
    Set xFg = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
    
LaCague:
    Me.MousePointer = vbDefault
    MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    Err.Clear
End Sub



Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        'If ButtonMenu.Index = 1 Then Modificar
        If ButtonMenu.Index = 2 Then CambiarEstado True
    End If
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Eliminar
        If ButtonMenu.Index = 2 Then CambiarEstado False
    End If
End Sub


'*****************************************************************************************************
'* Nombre Archivo   : CambiarEstado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL ESTADO DE UN REGISTRO DE LA TABLA ges_planaba
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Activado     |  Boolean   |  INDICA SI SE ACTIVA O DESACTIVA UN REGISTRO
'* DEVUELVE         :
'*****************************************************************************************************
Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar el plan de abastecimiento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar el plan de abastecimiento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    End If
    
    If Rpta = vbYes Then
        If Activado = False Then
            xCon.Execute "UPDATE ges_planaba SET ges_planaba.activo = 0 Where (((ges_planaba.id) = " & RstPlanAbas("id") & "))"
            MsgBox "El plan de abastecimiento se desactivo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            xCon.Execute "UPDATE ges_planaba SET ges_planaba.activo = -1 Where (((ges_planaba.id) = " & RstPlanAbas("id") & "))"
            MsgBox "El plan de abastecimiento se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    RstPlanAbas.Requery
    Dg1.Refresh
End Sub
