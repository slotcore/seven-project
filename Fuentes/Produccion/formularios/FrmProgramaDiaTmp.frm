VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmProgramaDiaTmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Programar Producción del Dia"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   5
      Top             =   360
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   -12465
         TabIndex        =   6
         Top             =   375
         Width           =   11820
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   45
            TabIndex        =   16
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
            Columns(1).Caption=   "Fch Trabajo"
            Columns(1).DataField=   "fchprod"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Programado Por"
            Columns(2).DataField=   "programador"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2117"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2037"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7673"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7594"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
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
            TabIndex        =   15
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Programa de Producción del Dia"
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
            Index           =   0
            Left            =   105
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   30
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6825
         Left            =   45
         TabIndex        =   9
         Top             =   375
         Width           =   11820
         Begin VB.CommandButton Cmd 
            Caption         =   "&Cargar de Programación Semanal"
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   3840
            TabIndex        =   23
            ToolTipText     =   "Agregar "
            Top             =   6465
            Width           =   2985
         End
         Begin VB.Frame Frame4 
            Caption         =   "( Periodo )"
            Height          =   615
            Left            =   9900
            TabIndex        =   21
            Top             =   0
            Width           =   1740
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
               TabIndex        =   22
               Top             =   315
               Width           =   1605
            End
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   2265
            Picture         =   "FrmProgramaDiaTmp.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   720
            Width           =   225
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Agregar"
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   3
            ToolTipText     =   "Agregar "
            Top             =   6465
            Width           =   1275
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1380
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar "
            Top             =   6465
            Width           =   1275
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   7260
            TabIndex        =   13
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   255
            Visible         =   0   'False
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   0
            Left            =   1305
            TabIndex        =   0
            Top             =   360
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
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5370
            Left            =   60
            TabIndex        =   2
            Top             =   1050
            Width           =   11700
            _cx             =   20637
            _cy             =   9472
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
            Rows            =   1
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmProgramaDiaTmp.frx":0132
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
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1305
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(0)"
            ToolTipText     =   "Ingrese DNI del Supervisor"
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Programado Por"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   19
            Top             =   795
            Width           =   1140
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(0)"
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
            Index           =   0
            Left            =   3510
            TabIndex        =   18
            Top             =   690
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   6705
            TabIndex        =   14
            Top             =   375
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fch Producción"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   465
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Programa la Producción del Dia"
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
            TabIndex        =   10
            Top             =   15
            Width           =   11610
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(0)"
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
            Index           =   0
            Left            =   2505
            TabIndex        =   20
            Top             =   690
            Width           =   3015
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
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
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":01C8
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":070C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":0A9E
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":0C22
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":1076
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":118E
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":16D2
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":1C16
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":1D2A
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":1E3E
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":2292
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDiaTmp.frx":23FE
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmProgramaDiaTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim Agregando As Boolean
Dim SeEjecuto As Boolean
Dim RstFrm As New ADODB.Recordset
'----
Dim mMesActivo  As Integer              '--
Public RstGrupoDet As New ADODB.Recordset  '--
Dim fOrdenLista As Boolean '--especfica el orden de la lista de la consulta
Dim mRowAdd As Double '--identificador unico por fila cuando se agrege una tarea

Dim mIdRegistro& '--identificador del registro

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd
        Case 1 '--eliminar
            pRegistroDel
        Case 2 '--cargar de programacion semanal
            pCargarProgramaSemanal
            
    End Select
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
'            If Index = 0 Then PopupMenu Menu3
'            If Index = 1 Then PopupMenu menu2
        End If
    End If

End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 1 And Col <> 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLNotId As String
    Dim nTitulo As String

    Select Case Col
        Case 1 '--producto
                ReDim xCampos(4, 4) As String
                xCampos(0, 0) = "Cod.Pro":    xCampos(0, 1) = "codpro":       xCampos(0, 2) = "2000":      xCampos(0, 3) = "C":
                xCampos(1, 0) = "Desripcion": xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4000":      xCampos(1, 3) = "C":
                xCampos(2, 0) = "Receta":     xCampos(2, 1) = "codrec":       xCampos(2, 2) = "1200":       xCampos(2, 3) = "C":
                xCampos(3, 0) = "Cant.Rec":   xCampos(3, 1) = "totrec":       xCampos(3, 2) = "1000":       xCampos(3, 3) = "N":
                        
                nSQLNotId = GRID_GENERAR_SQL_ID(Fg1, 3, " and alm_inventario.id ", "NOT IN", True)
                '--alm_inventario.activo =-1 and  SE MOSTRARAN TODOS
                nSQL = "SELECT id as iditem,idrec,codpro,descripcion,codrec,descrec ,totrec " _
                    + vbCr + " From " _
                    + vbCr + " (SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion , Count(pro_receta.codrec) AS totrec FROM alm_inventario LEFT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem WHERE alm_inventario.tippro=3 " & nSQLNotId & " GROUP BY alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion ) as item " _
                    + vbCr + " Left Join " _
                    + vbCr + " (SELECT pro_receta.id as idrec,pro_receta.iditem, pro_receta.codrec, pro_receta.descripcion AS descrec FROM pro_receta WHERE (((pro_receta.prirec)=1))) AS receta " _
                    + vbCr + " ON  item.id= receta.iditem"

                
                nTitulo = "Buscando Productos"
                            
        Case 2 '--receta
            If NulosN(Fg1.TextMatrix(Row, 3)) = 0 Then '--
                MsgBox "Seleccione el Producto", vbExclamation, xTitulo
                Fg1.Col = 1
                Fg1.SetFocus
                Exit Sub
            End If
            
                ReDim xCampos(2, 4) As String
                xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codrec":      xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion": xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    
                nSQL = "SELECT pro_receta.id as idrec, pro_receta.iditem, pro_receta.codrec, pro_receta.descripcion " _
                    + vbCr + " From pro_receta " _
                    + vbCr + " WHERE (((pro_receta.iditem)=" & NulosN(Fg1.TextMatrix(Row, 3)) & "));"


                nTitulo = "Buscando Receta"
                
    End Select

    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio, ""

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    
    If Col = 1 Then '--producto
        Fg1.TextMatrix(Row, 1) = NulosC(RstTmp.Fields("descripcion"))
        Fg1.TextMatrix(Row, 2) = NulosC(RstTmp.Fields("codrec"))
        Fg1.TextMatrix(Row, 3) = NulosN(RstTmp.Fields("iditem"))
        Fg1.TextMatrix(Row, 4) = NulosC(RstTmp.Fields("idrec"))
    ElseIf Col = 2 Then '--receta
        Fg1.TextMatrix(Row, 2) = NulosC(RstTmp.Fields("codrec"))
        Fg1.TextMatrix(Row, 4) = NulosC(RstTmp.Fields("idrec"))
        '-------------------------------------------------------------------
    End If
        
    Agregando = False
    Set RstTmp = Nothing
    Exit Sub
SALIR:
    Set RstTmp = Nothing
    Agregando = False
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col <= 10 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        cmd_Click 0
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        cmd_Click 1  'F4 = Eliminar Item
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Fg1_KeyUp"
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
'        If QueHace = 3 Then
'            PopupMenu Menu4
'        Else
'            PopupMenu Menu1
'        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    mRowAdd = -999
    mMesActivo = xMes
    pCargarGrid
    pConfigurarGrilla
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ningún Programa de Producción del Dia, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
    
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
   
    Dg3.Columns("fchprod").NumberFormat = FORMAT_DATE
    
    TabOne1.CurrTab = 0
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Dg3.HeadLines = 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    Else
        Set RstGrupoDet = Nothing
        Set RstFrm = Nothing
    End If
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
                If RstFrm.EOF = True Or RstFrm.BOF = True Then RstFrm.MoveFirst
            End If
            Dg3.Refresh
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then RstFrm.Filter = ""
    
    If Button.Index = 10 Then CambiarMes
    
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    TabOne1.CurrTab = 0
    If MsgBox("¿Esta seguro de eliminar la Programación de Producción?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETe * FROM pro_progdiadet WHERE idprogra = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM pro_progdia WHERE id = " & RstFrm("id") & ""
        
        MsgBox "La Producción del dia " + Format(RstFrm("fchprod"), "dd/mm/yy") + " fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
        
        RstFrm.Requery
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No se ha registrado ningún Programa de Producción del Dia, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
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

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    pHabilitarObj False
    Label1.Caption = "Detalle de la Programación de la Producción del Dia"
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    Dg3.SetFocus
End Sub

Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If

    QueHace = 2
    
    TabOne1.TabEnabled(0) = False
    ActivaTool
    pHabilitarObj True
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    Fg1.SelectionMode = flexSelectionFree
    
    Label1.Caption = "Modificando la Programación de la Producción del Dia"
    
    TxtFecha(0).Enabled = False
    TxtFecha(0).SetFocus
End Sub

Private Sub MuestraSegundoTab()
    With RstFrm
        
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
        If IsDate(.Fields("fchprod")) = True Then
            TxtFecha(0).valor = CDate(.Fields("fchprod"))
        End If
        If NulosN(.Fields("numdoc")) <> 0 Then
            txt_cb(0).Text = .Fields("numdoc")
            txt_cb_Validate 0, False
        End If

        MuestraDetalle
        
    End With
End Sub

Private Sub MuestraDetalle()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    
    '--------------------------------
    nSQL = "SELECT alm_inventario.descripcion, pro_receta.codrec, pro_progdiadet.canpro, pro_progdiadet.idund, pro_progdiadet.iditem, pro_progdiadet.idrec " _
        + vbCr + " FROM (alm_inventario INNER JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) INNER JOIN pro_progdiadet ON pro_receta.id = pro_progdiadet.idrec " _
        + vbCr + " WHERE (((pro_progdiadet.idprogra)=" & RstFrm("id") & "));"

    RST_Busq RstTmp, nSQL, xCon
    Agregando = True
    With Fg1
        .Rows = 1
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            DoEvents
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("descripcion"))
            .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("codrec"))
            .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("iditem"))
            .TextMatrix(.Rows - 1, 4) = NulosN(RstTmp.Fields("idrec"))
            '---
            RstTmp.MoveNext
        Loop
    End With
    '--------------------------------------------
    Set RstTmp = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub
error:
    SHOW_ERROR Me.Name, "MuestraDetalle"
    Me.MousePointer = vbDefault
    Set RstTmp = Nothing
    Agregando = False
End Sub

Private Sub pHabilitarObj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
    habilitar Cmd, band
End Sub

Sub Blanquea()
    LimpiaText TxtFecha
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    
    TabOne1.TabEnabled(0) = False
    ActivaTool
    If mMesActivo <> 0 And mMesActivo <> 13 Then TxtFecha(0).valor = CDate("01/" & mMesActivo & "/" & AnoTra)
    Blanquea
    pHabilitarObj True
    Label1.Caption = "Agregando el Programa de Producción del Dia"
    
    TxtFecha(0).Enabled = True
    TxtFecha(0).SetFocus
    pConfigurarGrilla
    Fg1.SelectionMode = flexSelectionFree
    '--agregando un registro por defecto
    Fg1.Rows = 2
    
    '--
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pImprimir True

    If ButtonMenu.Index = 2 Then pImprimir

End Sub

Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Programación de Producción del Dia " & TxtFecha(0).valor, vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xCod&, xCol&, xFil&
    
    On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_progdia ", xCon
        xCod = HallaCodigoTabla("pro_progdia", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        xCod = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pro_progdia WHERE id =" & xCod & "", xCon
        xCon.Execute "DELETE * FROM pro_progdiadet WHERE idprogra = " & xCod & ""
    
    End If
    
    mIdRegistro = xCod
        
    RST_Busq RstDet, "SELECT top 1 * FROM pro_progdiadet", xCon
    
    RstCab("fchprod") = CDate(TxtFecha(0).valor)
    RstCab("fchfin") = Null
    RstCab("idprog") = NulosN(lbl_cod(0).Caption)
    
    RstCab.Update
    
    For xFil = 1 To Fg1.Rows - 1
        RstDet.AddNew
        '--codigo
        RstDet("idprogra") = xCod
        RstDet("corr") = xFil
        '---
        RstDet("tipo") = 0
        RstDet("idref") = 0
        RstDet("canori") = 0
        RstDet("idundori") = 0
        RstDet("idrec") = NulosN(Fg1.TextMatrix(xFil, 4))
        RstDet("iditem") = NulosN(Fg1.TextMatrix(xFil, 3))
        RstDet("canpro") = 0
        RstDet("idund") = 0
        RstDet("idres") = 0
        RstDet("numlote") = 0
        RstDet.Update
        
    Next xFil
    
    xCon.CommitTrans
    MsgBox "La Programación de Producción del Dia se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    Grabar = True
SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

Private Function fValidarDatos() As Boolean
    If TxtFecha(0).valor = "" Or IsDate(TxtFecha(0).valor) = False Then
        MsgBox "No ha especificado la fecha de la Producción ", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    Dim band As Integer
    band = Validar(txt_cb)
'    If band <> -1 Then
'       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
'       txt_cb(band).SetFocus
'       Exit Function
'    End If

    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado el registro de las tareas", vbInformation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    '---------------------------------------------------------------------------
    '--validar la grilla
    Dim mRow&, mCol&
    
    mCol = -1
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 3)) = 0 Then '--producto
            MsgBox "Seleccione el Producto", vbExclamation, xTitulo
            mCol = 1
        ElseIf NulosN(Fg1.TextMatrix(mRow, 4)) = 0 Then '--receta
            MsgBox "Seleccione la Receta" & vbCr & "Producto: " & Fg1.TextMatrix(Fg1.Rows - 1, 1), vbExclamation, xTitulo
            mCol = 2
        End If

        If mCol <> -1 Then Exit For
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
        Fg1.SetFocus
        Exit Function
    End If
    '---------------------------------------------------------------------------

    fValidarDatos = True
End Function

Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL  As String
    
    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = lblperiodo(0).Caption
    
    nSQL = "SELECT pro_progdia.id, pro_progdia.fchprod, pro_emp.id AS idprog, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS programador, pla_empleados.numdoc " _
        + vbCr + " FROM pro_progdia LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_progdia.idprog = pro_emp.id " _
        + vbCr + " WHERE (((Month([pro_progdia].[fchprod]))=" & mMesActivo & ") AND ((Year([pro_progdia].[fchprod]))=" & AnoTra & ")) ORDER BY pro_progdia.fchprod ASC ;"
    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    TabOne1.CurrTab = 0
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Selecione un Periodo Correcto", vbExclamation, xTitulo
        CambiarMes
        Exit Sub
    End If
    pCargarGrid
End Sub

Private Sub pImprimir(Optional IMP_LISTADO As Boolean = False)

    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
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

Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Fch.Prod":         xCampos(0, 1) = "dia":       xCampos(0, 2) = "850":    xCampos(0, 3) = "F"
    xCampos(1, 0) = "N°.Prod":          xCampos(1, 1) = "num":       xCampos(1, 2) = "900":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Receta":           xCampos(2, 1) = "codrec":    xCampos(2, 2) = "1000":   xCampos(2, 3) = "C"
    xCampos(3, 0) = "Producto":         xCampos(3, 1) = "proddesc":  xCampos(3, 2) = "3200":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cantidad":         xCampos(4, 1) = "cantidad":  xCampos(4, 2) = "800":    xCampos(4, 3) = "N"
    xCampos(5, 0) = "Responsable":      xCampos(5, 1) = "resnom":    xCampos(5, 2) = "1500":   xCampos(5, 3) = "C"
        
        
    nSQL = " SELECT pro_produccion.id, pro_produccion.num, format(pro_produccion.dia,'dd/mm/yy') as dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, pla_empleados_2.ape & ' ' & pla_empleados_2.nom AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec " _
    + vbCr + " FROM ((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN ((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) ON pro_produccion.id = pro_producciondet.idpro " _
    + vbCr + " WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND MONTH(pro_produccion.dia)= " + CStr(mMesActivo) + ""

    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Producción", "dia", "proddesc", Principio
    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True And RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(RstTmp("id"))
SALIR:
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

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


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--programador
            nTitulo = "Buscando Personal"
            nSQL = "SELECT pro_emp.idemp, pro_emp.id AS idper, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, mae_dociden.abrev, pla_empleados.numdoc " _
                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
                + vbCr + " WHERE (((pro_empdet.idfun)=2)); "

    End Select
    
    ReDim xCampos(3, 3) As String
    xCampos(0, 0) = "Tipo Doc":             xCampos(0, 1) = "abrev":    xCampos(0, 2) = "850":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Número":               xCampos(1, 1) = "numdoc":   xCampos(1, 2) = "1200":  xCampos(1, 3) = "C"
    xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":  xCampos(2, 2) = "4500":  xCampos(1, 3) = "C"
    
            
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = NulosC(RstTmp.Fields("numdoc")) '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields("nombre")) '--NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields("idper"))  '--CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields("nombre")) '--NOMBRE
    
    Fg1.SetFocus
    
SALIR:
    Set RstTmp = Nothing
Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub


Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--personal
        
            nSQL = "SELECT pro_emp.idemp, pro_emp.id AS idper, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, mae_dociden.abrev, pla_empleados.numdoc " _
                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
                + vbCr + " WHERE (((pro_empdet.idfun)=2)) and pla_empleados.numdoc = '" & NulosC(txt_cb(0).Text) & "' "
    
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index) = NulosC(RstTmp.Fields("numdoc")) '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields("nombre")) '--NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields("idper"))  '--CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields("nombre")) '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
SALIR:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

'****************************************************************************************

Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > Fg1.FixedRows Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 3)) = 0 Then '--tipo de persona
            MsgBox "Seleccione el Producto", vbExclamation, xTitulo
            mCol = 1
        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 4)) = 0 Then '--persona/grupo
            MsgBox "Seleccione la Receta" & vbCr & "Producto: " & Fg1.TextMatrix(Fg1.Rows - 1, 1), vbExclamation, xTitulo
            mCol = 2
        Else
            fInsertar = True
            mCol = 1
        End If
    Else
        fInsertar = True
        mCol = 1
    End If
    
    If fInsertar = True Then Fg1.AddItem ""
    
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = mCol
    
    '--cargar el buscador por defecto
    If fInsertar = True Then Fg1_CellButtonClick Fg1.Rows - 1, 1

    Fg1.SetFocus
    Agregando = False

End Sub

Private Sub pRegistroDel()
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
    
    If Fg1.Rows > 1 Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 1
        Fg1.SetFocus
    Else
        Cmd(0).SetFocus
    End If
End Sub


Private Sub pConfigurarGrilla()
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
   
    With Fg1 '--de los ingredientes
        .Rows = 1
        .Cols = 5
        .FixedRows = 1
        .RowHeight(0) = 250
                
        .TextMatrix(0, 1) = "Descripción":  .ColWidth(1) = 5500:  .ColAlignment(1) = flexAlignLeftCenter:    .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Receta":       .ColWidth(2) = 2500:  .ColAlignment(2) = flexAlignLeftCenter:    .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "IdItem":       .ColWidth(3) = 0:
        .TextMatrix(0, 4) = "IdRec":        .ColWidth(4) = 0:
        
        .SelectionMode = flexSelectionByRow
                
        GRID_COMBOLIST Fg1, 1  '--producto
        GRID_COMBOLIST Fg1, 2 '--receta
        DoEvents
                
    End With

End Sub



Private Sub pCargarProgramaSemanal()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    If IsDate(TxtFecha(0).valor) = False Then
        MsgBox "Ingrese la Fecha de Producción", vbExclamation, xTitulo
        Exit Sub
    End If
    
    '--------------------------------
    nSQL = "SELECT alm_inventario.descripcion, pro_receta.codrec, pro_programadet.iditem, pro_programadet.idrec " _
        + vbCr + " FROM (pro_receta RIGHT JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) LEFT JOIN alm_inventario ON pro_programadet.iditem = alm_inventario.id " _
        + vbCr + " WHERE (((pro_programadet.dia)=CDate('" & TxtFecha(0).valor & "')));"

    RST_Busq RstTmp, nSQL, xCon
    DoEvents
    If RstTmp.RecordCount = 0 Then
        MsgBox "No hay Programación de la Semana para este Día", vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Sub
    Else
        If Fg1.Rows > 1 Then
            If MsgBox("Se procederá a Reiniciar la Lista Seleccionada" + vbCr + "Desea Continuar...", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Sub
        End If
    End If
    
    DoEvents
    Agregando = True
    With Fg1
        .Rows = 1
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            DoEvents
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("descripcion"))
            .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("codrec"))
            .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("iditem"))
            .TextMatrix(.Rows - 1, 4) = NulosN(RstTmp.Fields("idrec"))
            '---
            RstTmp.MoveNext
        Loop
    End With
    

    '--------------------------------------------
    Set RstTmp = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub
error:
    SHOW_ERROR Me.Name, "pCargarProgramaSemanal"
    Me.MousePointer = vbDefault
    Set RstTmp = Nothing
    Agregando = False
End Sub


