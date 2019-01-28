VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmMantActualizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEVEN - Actualizador"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "FrmMantActualizar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmRuta 
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   10050
      TabIndex        =   7
      Top             =   2310
      Visible         =   0   'False
      Width           =   3390
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   420
         Left            =   1770
         TabIndex        =   12
         Top             =   2160
         Width           =   1020
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   420
         Left            =   630
         TabIndex        =   11
         Top             =   2160
         Width           =   1020
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   60
         TabIndex        =   10
         Top             =   735
         Width           =   3240
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   390
         Width           =   3240
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3120
         Picture         =   "FrmMantActualizar.frx":030A
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   8
         ToolTipText     =   "Cerrar"
         Top             =   75
         Width           =   195
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar Directorio"
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
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Width           =   1905
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   3375
         X2              =   3375
         Y1              =   15
         Y2              =   3465
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -330
         X2              =   3715
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   -30
         X2              =   4015
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   15
         Y2              =   3450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   0
         X2              =   3500
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   30
         Top             =   45
         Width           =   3315
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5220
      Left            =   0
      TabIndex        =   0
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4800
         Left            =   -10425
         TabIndex        =   3
         Top             =   375
         Width           =   9780
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4425
            Left            =   60
            TabIndex        =   6
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
            Columns(1).Caption=   "Fch. Reg."
            Columns(1).DataField=   "xfchreg"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Ruta"
            Columns(3).DataField=   "origen"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tipo"
            Columns(4).DataField=   "xtipo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Estado"
            Columns(5).DataField=   "xestado"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2011"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1931"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=3096"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3016"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=5503"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5424"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1799"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1720"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=78,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consultando Actualizaciones"
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
            TabIndex        =   5
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
            TabIndex        =   4
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4800
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   9780
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   4425
            Left            =   60
            TabIndex        =   14
            Top             =   330
            Width           =   9660
            _cx             =   17039
            _cy             =   7805
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
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "  Datos Principales  |  PC por Actualizar  |  PC Actualizadas  "
            Align           =   0
            CurrTab         =   0
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Height          =   4005
               Left            =   45
               TabIndex        =   20
               Top             =   45
               Width           =   9570
               Begin VB.ComboBox CbTipo 
                  Height          =   315
                  ItemData        =   "FrmMantActualizar.frx":05F6
                  Left            =   2040
                  List            =   "FrmMantActualizar.frx":0600
                  Style           =   2  'Dropdown List
                  TabIndex        =   25
                  Top             =   1260
                  Width           =   2715
               End
               Begin VB.CommandButton CmdBusArch 
                  Height          =   240
                  Left            =   7765
                  Picture         =   "FrmMantActualizar.frx":061A
                  Style           =   1  'Graphical
                  TabIndex        =   24
                  Top             =   960
                  Width           =   240
               End
               Begin VB.TextBox TxtDescripcion 
                  Height          =   315
                  Left            =   2040
                  Locked          =   -1  'True
                  TabIndex        =   23
                  Text            =   "TxtDescripcion"
                  Top             =   570
                  Width           =   6000
               End
               Begin VB.TextBox TxtGlosa 
                  Height          =   630
                  Left            =   2040
                  Locked          =   -1  'True
                  MaxLength       =   11
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   22
                  Text            =   "FrmMantActualizar.frx":074C
                  Top             =   1650
                  Width           =   6000
               End
               Begin VB.TextBox TxtRuta 
                  Height          =   285
                  Left            =   2040
                  Locked          =   -1  'True
                  TabIndex        =   21
                  Text            =   "TxtRuta"
                  Top             =   930
                  Width           =   6000
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
                  Height          =   300
                  Left            =   2040
                  TabIndex        =   26
                  Top             =   225
                  Width           =   1275
                  _ExtentX        =   2249
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
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha"
                  Height          =   195
                  Index           =   0
                  Left            =   390
                  TabIndex        =   31
                  Top             =   330
                  Width           =   450
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Glosa"
                  Height          =   195
                  Index           =   7
                  Left            =   390
                  TabIndex        =   30
                  Top             =   1740
                  Width           =   405
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Ruta"
                  Height          =   195
                  Index           =   2
                  Left            =   390
                  TabIndex        =   29
                  Top             =   1020
                  Width           =   345
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripción"
                  Height          =   225
                  Index           =   6
                  Left            =   390
                  TabIndex        =   28
                  Top             =   660
                  Width           =   840
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo"
                  Height          =   195
                  Index           =   10
                  Left            =   390
                  TabIndex        =   27
                  Top             =   1380
                  Width           =   315
               End
            End
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   4005
               Left            =   10305
               TabIndex        =   15
               Top             =   45
               Width           =   9570
               Begin VB.Frame Frame5 
                  Height          =   3600
                  Left            =   7980
                  TabIndex        =   16
                  Top             =   30
                  Width           =   1560
                  Begin VB.CommandButton CmdAdd 
                     Caption         =   "Agregar "
                     Height          =   690
                     Left            =   150
                     TabIndex        =   18
                     Top             =   1020
                     Width           =   1305
                  End
                  Begin VB.CommandButton CmdDel 
                     Caption         =   "Eliminar"
                     Height          =   690
                     Left            =   150
                     TabIndex        =   17
                     Top             =   1935
                     Width           =   1305
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   3510
                  Left            =   30
                  TabIndex        =   19
                  Top             =   120
                  Width           =   7710
                  _cx             =   13600
                  _cy             =   6191
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
                  Rows            =   50
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmMantActualizar.frx":0755
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
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   4005
               Left            =   10605
               TabIndex        =   32
               Top             =   45
               Width           =   9570
               _cx             =   16880
               _cy             =   7064
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
               Rows            =   50
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmMantActualizar.frx":07B1
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Actualización"
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
            TabIndex        =   2
            Top             =   30
            Width           =   9675
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7065
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
            Picture         =   "FrmMantActualizar.frx":085E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":0DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":1134
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":12B8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":170C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":1824
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":1D68
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":22AC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":23C0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":24D4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":2928
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":2A94
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantActualizar.frx":2FDC
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
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
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar "
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
                  Text            =   "Eliminar "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar "
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Tag             =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmMantActualizar"
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
Dim Agregando As Boolean
Dim mIdRegistro& '--identificador del registro


'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long
'------------


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
    
    TxtFecha.Valor = NulosC(RstEmp("fchreg"))
    TxtDescripcion.Text = NulosC(RstEmp("descripcion"))
    TxtRuta.Text = NulosC(RstEmp("origen"))
    CbTipo.ListIndex = NulosN(RstEmp("tipo"))
    TxtGlosa.Text = NulosC(RstEmp("glosa"))
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    '--Cargar Detalle de pc's x actualizar
    Dim nSQL As String
    Dim xRs As New ADODB.Recordset
    nSQL = "SELECT mae_pc.id, mae_pc.pc " _
        & " FROM mae_versiondet INNER JOIN mae_pc ON mae_versiondet.idpc = mae_pc.id " _
        & " WHERE (((mae_versiondet.idver)=" & NulosN(RstEmp("id")) & ")); "
    RST_Busq xRs, nSQL, xConRuta
    
    If xRs.RecordCount <> 0 Then
    
        xRs.MoveFirst
        Do While Not xRs.EOF
            Agregando = True
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("id"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("pc"))
            
            xRs.MoveNext
        Loop
        
    End If
    Set xRs = Nothing
    
    '--Cargar Detalle de pc's que se actualizaron
    nSQL = "SELECT mae_pc.id, mae_pc.pc, mae_versionact.userpc, mae_versionact.fchact, mae_versionact.horact " _
        & " FROM mae_pc RIGHT JOIN mae_versionact ON mae_pc.id = mae_versionact.idpc " _
        & " Where (((mae_versionact.idver) = " & NulosN(RstEmp("id")) & ")) " _
        & " ORDER BY mae_versionact.fchact, mae_versionact.horact "
        
    RST_Busq xRs, nSQL, xConRuta
    If xRs.RecordCount <> 0 Then
    
        xRs.MoveFirst
        Do While Not xRs.EOF
            Agregando = True
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(xRs("pc"))
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(xRs("userpc"))
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(xRs("fchact"))
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(xRs("horact"))
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosC(xRs("id"))
            xRs.MoveNext
        Loop
        
    End If
    
    Set xRs = Nothing
    
    
End Sub

Private Sub CbTipo_Validate(Cancel As Boolean)
    If CbTipo.ListIndex = 1 Then
        TabOne2.TabEnabled(1) = True
    Else
        TabOne2.TabEnabled(1) = False
    End If
End Sub

Private Sub CmdAdd_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLIdPc As String
      
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Id Pc":        xCampos(0, 1) = "id":       xCampos(0, 2) = "800":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nombre Pc":    xCampos(1, 1) = "pc":       xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
       
    nSQLIdPc = GRID_GENERAR_SQL_ID(Fg1, 1, " where mae_pc.id", " NOT IN ", True)
    
    nSQL = "SELECT 0 as xsel, mae_pc.id, mae_pc.pc FROM mae_pc " & nSQLIdPc
            
    CARGAR_DLL_EPSBUSCAR_SEL xConRuta, xRs, nSQL, xCampos(), "Buscando Pc"
    
    If xRs.State = 0 Then GoTo xSalir
    If xRs.RecordCount = 0 Then GoTo xSalir

    Do While Not xRs.EOF
        Agregando = True
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("pc"))
        
        xRs.MoveNext
    Loop
    
xSalir:
    Set xRs = Nothing

End Sub

Private Sub CmdBusArch_Click()
    If QueHace = 3 Then Exit Sub
    FrmRuta.Visible = True
    FrmRuta.Top = 1680
    FrmRuta.Left = 3270
    Drive1.SetFocus
End Sub

Private Sub CmdAceptar_Click()
    FrmRuta.Visible = False
    
    If NulosC(TxtRuta.Text) <> "" Then
        If Right(TxtRuta.Text, 1) <> "\" Then
            TxtRuta.Text = TxtRuta.Text & "\"
        End If
    End If
    
End Sub

Private Sub CmdCancelar_Click()
    FrmRuta.Visible = False
    TxtRuta.Text = ""
    TxtRuta.SetFocus
End Sub

Private Sub CmdDel_Click()

    If QueHace = 3 Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No hay registros para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
    
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
        IdMenuActivo = 252
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
                
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
        RST_Busq RstEmp, "SELECT mae_version.*, [mae_version].[fchreg] & '' AS xfchreg, IIf(mae_version.estado=0,'Cerrado','Abierto') AS xestado, IIf([mae_version].[tipo]=0,'Todos','Personalizado') AS xtipo FROM mae_version order by mae_version.fchreg desc ", xConRuta

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
    Label5.Caption = "Agregando Actualización"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaToolbar
    Blanquea
    Bloquea True
    
    TxtFecha.Valor = Date
    TabOne2.CurrTab = 0
    CbTipo.ListIndex = 0
    TxtFecha.SetFocus
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
    Label5.Caption = "Modificando Actualizador"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaToolbar
    Blanquea
    Bloquea True
    MuestraSegundoTab
    TabOne2.CurrTab = 0
    TxtFecha.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xId As Long
    TabOne1.CurrTab = 0
    If RstEmp.EOF = True Or RstEmp.BOF = True Or RstEmp.RecordCount = 0 Then
        MsgBox "No hay Registros para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    xId = RstEmp("id")
    
    Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xConRuta.Execute "DELETE * FROM mae_versionact WHERE idver = " & xId & ""
        xConRuta.Execute "DELETE * FROM mae_versiondet WHERE idver = " & xId & ""
        xConRuta.Execute "DELETE * FROM mae_version WHERE id = " & xId & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
                
        RstEmp.Requery
        Dg1.Refresh
        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTARA AL CARGAR EL EVENTO LOAD
    SeEjecuto = False
    QueHace = 3
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    Fg2.ColWidth(5) = 0
    
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstEmp.Requery
            Dg1.Refresh
            
            If RstEmp.RecordCount <> 0 Then
                RstEmp.MoveFirst
                RstEmp.Find "id=" & mIdRegistro
                If RstEmp.EOF = True Then RstEmp.MoveFirst
            End If
            
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then TDB_Actualizar Me, TabOne1, Dg1, RstEmp
    
    If Button.Index = 14 Then
        Set RstEmp = Nothing
        Unload Me
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then Modificar
        If ButtonMenu.Index = 2 Then Activar
    End If
    
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Eliminar
        If ButtonMenu.Index = 2 Then Anular
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
    If NulosC(TxtFecha.Valor) = "" Then
        MsgBox "No ha especificado la Fecha", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFecha.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtDescripcion.Text) = "" Then
        MsgBox "No ha especificado la descripción de la actualización", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtRuta.Text) = "" Then
        MsgBox "No ha especificado la ruta de la actualización", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRuta.SetFocus
        Exit Function
    End If
    
    If CbTipo.ListIndex = -1 Then
        MsgBox "No ha especificado el tipo de actualización", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CbTipo.SetFocus
        Exit Function
    End If
    '--Verificar si ruta existe
    If QueHace = 1 Then
        If A.FolderExists(TxtRuta.Text) = False Then
            MsgBox "La ruta no existe" & vbCr & "Proceda a crear la carpeta manualmente, luego vuelva seleccionar", vbInformation, xTitulo
            TxtRuta.SetFocus
        End If
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim B As Integer
    
    On Error GoTo lAcAGUE
    xConRuta.BeginTrans
    
    ' GRAMAOS LOS DATOS
    If QueHace = 1 Then
        ' OBTENEMOS EL ID PARA EL NUEVO REGITROS
        xId = HallaCodigoTabla("mae_version", xConRuta, "id")
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM mae_version", xConRuta
        
        ' AGREGAMOS UN NUEVO REGISTRO
        RstCab.AddNew
        RstCab("id") = xId
    Else
        ' BUSCAMOS EL REGISTRO Y TRAEMOS LOS DATOS
        xId = RstEmp("id")
        RST_Busq RstCab, "SELECT * FROM mae_version WHERE id  = " & xId & "", xConRuta
        xConRuta.Execute "delete from mae_versiondet where idver=" & xId & ""
        
    End If
    
    mIdRegistro = xId
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM mae_versiondet", xConRuta
    
    ' ASIGNAMOS LOS DATOS A CADA CAMPO
    RstCab("fchreg") = CDate(TxtFecha.Valor)
    RstCab("descripcion") = TxtDescripcion.Text
    RstCab("glosa") = TxtGlosa.Text
    RstCab("origen") = TxtRuta.Text
    RstCab("estado") = 0
    RstCab("tipo") = CbTipo.ListIndex
    RstCab.Update
    
    '--GRABAMOS EL DETALLE
    For B = 1 To Fg1.Rows - 1
        
        RstDet.AddNew
        RstDet("idver") = xId
        RstDet("idpc") = Fg1.TextMatrix(B, 1)
        
        RstDet.Update
    Next B
    
    xConRuta.CommitTrans
    
    Grabar = True
    
    If QueHace = 1 Then
        MsgBox "El registro se generó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Else
        MsgBox "El registro se modificó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    Exit Function
    
lAcAGUE:
    Resume
    xConRuta.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
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
    Label5.Caption = "Detalle de la Actualización"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
End Sub


Private Sub TxtRuta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
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
    TxtFecha.Valor = ""
    TxtDescripcion.Text = ""
    TxtRuta.Text = ""
    CbTipo.ListIndex = -1
    TxtGlosa.Text = ""
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Bloquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Bloquea(xBand As Boolean)
    TxtFecha.Locked = Not xBand
    TxtDescripcion.Locked = Not xBand
    TxtRuta.Locked = Not xBand
    CbTipo.Locked = Not xBand
    TxtGlosa.Locked = Not xBand
End Sub

Private Sub Drive1_Change()
    On Error GoTo lAcAGUE
    Err.Clear
    Dir1.Path = Drive1
    Exit Sub
lAcAGUE:
    MsgBox Err.Description & vbCr, vbInformation, xTitulo
    Err.Clear
End Sub

Private Sub Dir1_Change()
    TxtRuta.Text = Dir1.Path
End Sub

Private Sub FrmRuta_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    OrigFX = x
    OrigFY = Y
    FrmRuta.ZOrder 0
End Sub

Private Sub FrmRuta_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 0 Then
        With FrmRuta
            .Move .Left + x - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub TxtRuta_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = vbKeyF5 Then CmdBusArch_Click
End Sub

Sub Activar()
    Dim Rpta As Integer
    If RstEmp.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbInformation, xTitulo
        Exit Sub
    End If
    If RstEmp("estado") = -1 Then
        MsgBox "El registro está activo", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de activar el registro?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
    
    
        mIdRegistro = NulosN(RstEmp("id"))
    
        xConRuta.Execute "UPDATE mae_version SET mae_version.estado = -1 WHERE (((mae_version.id)=" & NulosN(RstEmp("id")) & "))"
        
        'grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(RstEmp("id"))
        
        RstEmp.Requery
        
        If RstEmp.RecordCount <> 0 Then
            RstEmp.MoveFirst
            RstEmp.Find "id=" & mIdRegistro
            If RstEmp.EOF = True Then RstEmp.MoveFirst
        End If
        
        MsgBox "El registro se activó con éxito", vbInformation + vbOKOnly + vbDefaultButton1
        
        Exit Sub
    End If
End Sub

Sub Anular()
    Dim Rpta As Integer
    
    If RstEmp.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbInformation, xTitulo
        Exit Sub
    End If
    If RstEmp("estado") = 0 Then
        MsgBox "El registro está cerrado", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de cerrar el registro?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
    
        mIdRegistro = NulosN(RstEmp("id"))
        
        xConRuta.Execute "UPDATE mae_version SET mae_version.estado = 0 WHERE (((mae_version.id)=" & NulosN(RstEmp("id")) & "))"
        
        'grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(RstEmp("id"))

        RstEmp.Requery
        
        If RstEmp.RecordCount <> 0 Then
            RstEmp.MoveFirst
            RstEmp.Find "id=" & mIdRegistro
            If RstEmp.EOF = True Then RstEmp.MoveFirst
        End If
        
        MsgBox "El registro se cerró con éxito", vbInformation + vbOKOnly + vbDefaultButton1
        
        Exit Sub
    End If
End Sub
