VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmSelecciona1 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Registros"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid70.TDBGrid grilla 
      Height          =   4425
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   7805
      _LayoutType     =   4
      _RowHeight      =   13
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Id"
      Columns(0).DataField=   "id"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Sel"
      Columns(1).DataField=   "xsel"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   21
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   185
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=21"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=635"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=556"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=370"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=291"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=397"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=318"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=397"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=318"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=344"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=265"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=370"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=291"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=503"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=423"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=450"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=370"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(56)=   "Column(9).Width=370"
      Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=291"
      Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(62)=   "Column(10).Width=318"
      Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=238"
      Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(68)=   "Column(11).Width=344"
      Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=265"
      Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(72)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(74)=   "Column(12).Width=370"
      Splits(0)._ColumnProps(75)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(12)._WidthInPix=291"
      Splits(0)._ColumnProps(77)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(78)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(79)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(80)=   "Column(13).Width=370"
      Splits(0)._ColumnProps(81)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(13)._WidthInPix=291"
      Splits(0)._ColumnProps(83)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(84)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(85)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(86)=   "Column(14).Width=318"
      Splits(0)._ColumnProps(87)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(14)._WidthInPix=238"
      Splits(0)._ColumnProps(89)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(90)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(91)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(92)=   "Column(15).Width=370"
      Splits(0)._ColumnProps(93)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(15)._WidthInPix=291"
      Splits(0)._ColumnProps(95)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(96)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(97)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(98)=   "Column(16).Width=265"
      Splits(0)._ColumnProps(99)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(16)._WidthInPix=185"
      Splits(0)._ColumnProps(101)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(102)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(103)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(104)=   "Column(17).Width=450"
      Splits(0)._ColumnProps(105)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(17)._WidthInPix=370"
      Splits(0)._ColumnProps(107)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(108)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(109)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(110)=   "Column(18).Width=318"
      Splits(0)._ColumnProps(111)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(112)=   "Column(18)._WidthInPix=238"
      Splits(0)._ColumnProps(113)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(114)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(115)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(116)=   "Column(19).Width=423"
      Splits(0)._ColumnProps(117)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(118)=   "Column(19)._WidthInPix=344"
      Splits(0)._ColumnProps(119)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(120)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(121)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(122)=   "Column(20).Width=344"
      Splits(0)._ColumnProps(123)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(124)=   "Column(20)._WidthInPix=265"
      Splits(0)._ColumnProps(125)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(126)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(127)=   "Column(20).Order=21"
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
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1.5
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
      _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=106,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=118,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
      _StyleDefs(120) =   "Named:id=33:Normal"
      _StyleDefs(121) =   ":id=33,.parent=0"
      _StyleDefs(122) =   "Named:id=34:Heading"
      _StyleDefs(123) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(124) =   ":id=34,.wraptext=-1"
      _StyleDefs(125) =   "Named:id=35:Footing"
      _StyleDefs(126) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(127) =   "Named:id=36:Selected"
      _StyleDefs(128) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(129) =   "Named:id=37:Caption"
      _StyleDefs(130) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(131) =   "Named:id=38:HighlightRow"
      _StyleDefs(132) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(133) =   "Named:id=39:EvenRow"
      _StyleDefs(134) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(135) =   "Named:id=40:OddRow"
      _StyleDefs(136) =   ":id=40,.parent=33"
      _StyleDefs(137) =   "Named:id=41:RecordSelector"
      _StyleDefs(138) =   ":id=41,.parent=34"
      _StyleDefs(139) =   "Named:id=42:FilterBar"
      _StyleDefs(140) =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "( Opciones de Seleccion )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   60
      TabIndex        =   5
      Top             =   15
      Width           =   10875
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "Deseleccionar todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2775
         TabIndex        =   7
         Top             =   360
         Width           =   2040
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Seleccionar todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   6
         Top             =   360
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   765
      Left            =   30
      TabIndex        =   1
      Top             =   5130
      Width           =   10875
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "&Limpiar Filtro"
         Height          =   405
         Left            =   4800
         TabIndex        =   3
         Top             =   270
         Width           =   1155
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   6060
         TabIndex        =   4
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   3540
         TabIndex        =   2
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "&Activar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "&Desactivar"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Activar Todos Registros"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Desactivar Todos Registros"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_7 
         Caption         =   "Limpiar Filtro"
      End
   End
End
Attribute VB_Name = "FrmSelecciona1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Acepto As Boolean
Public Rst As New ADODB.Recordset
Dim SeEjecuto As Boolean

Private Sub CmdAceptar_Click()
    Acepto = True
    Me.Hide
    
    Rst.Filter = ""
    Rst.Filter = "xsel=-1"
    
    Me.Hide
End Sub


Private Sub CmdCancelar_Click()
    
    Acepto = False
    Unload Me
End Sub


Private Sub CmdLimpiar_Click()
    Menu1_7_Click
End Sub

Private Sub Form_Activate()
    Dim A As Integer
    
    If SeEjecuto = False Then
    
        
        SeEjecuto = True
        
CrearEncabezado

        ''''''''''''''''''
        '--agregando columna de seleccion al rst
        
        'xSQLCad = Replace(xSQLCad, "SELECT", "SELECT 0 as xsel,")
       
        Dim RstTemp As New ADODB.Recordset '--rsttemporal
        '--ejecutando consulta
        F_RST_Busq RstTemp, xSQLCad, xConeccion
        If RstTemp.RecordCount = 0 Then
            MsgBox "No se han encontrado registros con las condiciones especificadas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Acepto = False
            Unload Me
            Exit Sub
        End If
                
        '-----
                
        Dim obj As Object
        Set obj = CreateObject("SGI2_funciones.JC_Varios")
        
        Set Rst = Nothing
        
        '--definir estructura de rst
        'obj.DEFINIR_RST_TMP Rst, RstTemp
        '--se colocara como tipo de datos por defecto cadena para poder aplicar el filtro
        For A = 0 To RstTemp.Fields.Count - 1
            'Rst.Fields.Append RstTemp.Fields(a).Name,RstTemp.Fields(I).Type, -1, adFldIsNullable
            Rst.Fields.Append RstTemp.Fields(A).Name, adVarChar, -1, adFldIsNullable
        Next
        Rst.Open
        
        '--cargar datos a rst
        obj.CARGAR_RST_TMP Rst, RstTemp
        
        Set obj = Nothing
        
        Set grilla.DataSource = Rst
        
        '--posicionar en la primera fila
        If Rst.RecordCount <> 0 Then Rst.MoveFirst
        '-----------
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Acepto = False
    Unload Me
End Sub

Private Sub grilla_DblClick()
    On Error Resume Next
    If Rst.State = 0 Then Exit Sub
    If Rst.RecordCount < 1 Then Exit Sub
    Rst.Fields("xsel") = Not Rst.Fields("xsel")
    Err.Clear
End Sub

Private Sub grilla_FilterChange()
    Dim obj As Object
    Set obj = CreateObject("SGI2_funciones.JC_TDBGrid")
    obj.TDB_FiltroGenerar grilla, Rst
    Set obj = Nothing
End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grilla_DblClick
End Sub


Private Sub grilla_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu Menu1
End Sub
 
Private Sub Option1_Click()
    Menu1_4_Click
End Sub

Private Sub Option2_Click()
    Menu1_5_Click
End Sub


Private Sub Menu1_1_Click()
    Dim obj As Object
    Set obj = CreateObject("SGI2_funciones.JC_TDBGrid")
    obj.TDB_SelDesActCheck grilla, Rst, "xsel", "-1"
    Set obj = Nothing
End Sub

Private Sub Menu1_2_Click()
    Dim obj As Object
    Set obj = CreateObject("SGI2_funciones.JC_TDBGrid")
    obj.TDB_SelDesActCheck grilla, Rst, "xsel", "0"
    Set obj = Nothing
End Sub

Private Sub Menu1_4_Click()
    Dim obj As Object
    Set obj = CreateObject("SGI2_funciones.JC_TDBGrid")
    obj.TDB_TodosDesActCheck grilla, Rst, "xsel", "-1"
    Set obj = Nothing
End Sub

Private Sub Menu1_5_Click()
    Dim obj As Object
    Set obj = CreateObject("SGI2_funciones.JC_TDBGrid")
    obj.TDB_TodosDesActCheck grilla, Rst, "xsel", "0"
    Set obj = Nothing
End Sub

Private Sub Menu1_7_Click()
    '--limpiar los filtros
    Dim obj As Object
    Set obj = CreateObject("SGI2_funciones.JC_TDBGrid")
    obj.TDB_FiltroLimpiar grilla
    Rst.Filter = "xsel=-1"
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Do While Not Rst.EOF
            Rst("xsel") = "0"
            Rst.MoveNext
        Loop
    End If
    Rst.Filter = ""
        
    If Rst.RecordCount <> 0 Then Rst.MoveFirst
    
    Set obj = Nothing
End Sub


Sub CrearEncabezado()
    Dim A As Integer
    Dim B As Integer
    Dim xCol As Integer
    B = 1
    
    grilla.BatchUpdates = False
    
    xCol = 2 'grilla.Columns.Count
    
    For A = LBound(xCampos) To UBound(xCampos)
        grilla.Columns(xCol).DataField = xCampos(A, 1)  '"ruc"
        grilla.Columns(xCol).Caption = xCampos(A, 0) '"Nº. R.U.C."
        grilla.Columns(xCol).Width = xCampos(A, 2) '1500
        Select Case UCase(xCampos(A, 3))
            Case "C": grilla.Columns(xCol).Alignment = dbgLeft
            Case "D", "F": grilla.Columns(xCol).Alignment = dbgCenter
                grilla.Columns(xCol).NumberFormat = "dd/mm/yy"
            Case "N": grilla.Columns(xCol).Alignment = dbgRight
                grilla.Columns(xCol).NumberFormat = "###,###,##0.00"
            Case Else: grilla.Columns(xCol).Alignment = dbgLeft
        End Select
        xCol = xCol + 1
        
        If A = UBound(xCampos) - 1 Then
            Exit For
        End If
        
    Next A
    
    '--ocultando las columnas no utilizadas
    For A = UBound(xCampos) + 2 To grilla.Columns.Count - 1
        grilla.Columns(A).Visible = False
    Next A
    
End Sub

